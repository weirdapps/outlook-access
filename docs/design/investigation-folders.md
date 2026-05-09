# Investigation: Outlook CLI — Folder Management (search, create, move, list-in-folder)

Investigation date: 2026-04-21
Inputs consumed:

- `docs/design/refined-request-folders.md` (ground truth for scope, constraints, CLI
  shapes, error-mapping table)
- `docs/reference/codebase-scan-folders.md` (authoritative plug-in points — the
  existing `OutlookClient`, `ensureSession`/`mapHttpError`, `CliError` taxonomy,
  `ColumnSpec` output layer, and the 401-retry-once envelope)
- `docs/reference/codebase-scan-outlook-cli.md` (base scan: project conventions,
  TS strict mode, `launchPersistentContext`, `$HOME/.outlook-cli/` layout)
- `docs/research/outlook-v2-attachments.md` (calibration target for research
  depth — the reference for how much REST-v2 detail to pin down before build)
- `docs/design/investigation-outlook-cli.md` (mirrored section structure: Problem
  restatement → Approach comparison → Recommendation → Implementation strategy →
  Risk register → Research guidance)
- `src/http/outlook-client.ts` and `src/commands/list-mail.ts` (current code —
  read to confirm every assertion about where new code plugs in)

---

## 1. Problem restatement

The shipped `outlook-cli` reads mail / attachments / calendar from the signed-in
user's primary mailbox via the Outlook REST v2.0 surface
(`https://outlook.office.com/api/v2.0/me/...`), gated on a Playwright-captured
Bearer + cookie jar. We now need to extend it with folder-management
capabilities: search / list folders, resolve a folder by name or path, create a
folder (possibly nested), move a message (or a batch of messages) to a
destination folder, and allow `list-mail` to target any user-created folder —
not just the five well-known aliases it accepts today.

This is nontrivial for four reasons:

1. **Path semantics are a client invention.** Outlook REST v2.0 exposes folders
   as a tree via `ParentFolderId`; it does not accept a display-name path
   (`Inbox/Projects/Alpha`) directly. The CLI must resolve the path client-side
   with a per-level walk. The resolver owns the entire semantics of separators,
   escaping, case-folding, Unicode NFC, ambiguity, and well-known precedence.

2. **"Already exists" is a tri-state on the wire.** A `POST /childfolders` with
   a DisplayName that already exists under the target parent can return
   **409 Conflict** on most tenants and **400 Bad Request** on a few; the
   refined spec §7 flags this explicitly. The resolver must treat both as
   "collision" when `--idempotent` is set, otherwise let the original status
   surface. This is the single most fragile point of the folder create flow.

3. **`/move` returns a new id.** `POST /me/messages/{id}/move` **creates a copy
   in the destination and deletes the source**; the response body is the moved
   message with a **fresh id**. Scripts that chain `list-mail | move-mail`
   cannot re-use the source id. The CLI surface must expose the id mapping
   explicitly (`{sourceId, newId}` per moved message) or users will silently
   lose the ability to further address the moved item.

4. **Pagination for `/childfolders` is real.** A mailbox can have dozens of
   top-level folders and hundreds of descendants. The existing `OutlookClient`
   has no paging helper — every call issues one `GET` and returns one page.
   The folder enumeration paths must follow `@odata.nextLink` until the
   collection is drained (or a safety cap is hit). Adding that capability to
   the shared client avoids every folder command re-implementing it.

The constraints from `refined-request-folders.md` lock additional decisions:
project-wide error taxonomy is unchanged (no new exit codes 0/1/2/3/4/5/6 are
invented), no new mandatory config is introduced, concurrency is explicitly
serial (NG10), and no `$batch` is used (NG4) — every multi-message move is
N sequential REST calls.

---

## 2. Approach comparison

Each sub-section compares the realistic options for one folder sub-capability,
evaluated against the fixed constraints (single-mailbox, v2.0-only, Bearer-cookie
auth already owned by `OutlookClient`, no admin consent, no extra config).

### A. Folder lookup / path resolution

#### A1. Recursive walk — `GET /MailFolders/{parentId}/childfolders` per segment (RECOMMENDED)

For a path `Inbox/Projects/Alpha`:

1. Resolve segment 0 (`Inbox`) as a well-known alias or top-level
   DisplayName lookup.
2. `GET /me/MailFolders/{Inbox.Id}/childfolders?$top=100&$select=Id,DisplayName,...`
   — list direct children; locate `Projects` client-side (NFC + case-fold).
3. `GET /me/MailFolders/{Projects.Id}/childfolders?$top=100&$select=...` — list
   direct children; locate `Alpha`.
4. Return the resolved `Alpha` folder.

Client-side matching gives us the escaping, NFC, ambiguity, and
case-sensitivity story in one place. Pagination is handled by the shared
`listAll<T>` helper from §4.1. `@odata.nextLink` is followed; a cap of 50 pages
per level (refined §7) bounds the worst-case fan-out.

- **Pros**
  - Works reliably without needing OData `$filter` server-side support.
  - Gives us a natural spot to implement `--case-sensitive`, `--first-match`,
    `--include-hidden`, and ambiguity detection.
  - One HTTP shape per level; the paging helper is reused verbatim by
    `list-folders --recursive`.
  - Tolerant to tenants that return slightly different `$filter` semantics
    (see A2 below).
- **Cons**
  - Each path segment is one or more REST calls (depth N + Σ pages). A 3-level
    path with 120 siblings at level 2 is 2+1 calls — acceptable.
  - Full-children downloads are slightly wasteful when the target is known by
    name; `$select` keeps the payload tight (~120-200 bytes per folder).
- **Failure modes**
  - A tenant returns a page smaller than requested with no `@odata.nextLink` —
    normal OData behaviour (tenant server's own cap); the helper stops
    cleanly.
  - 429 on a recursive walk (`list-folders --recursive` across a huge tree) —
    the existing `OutlookClient` surfaces `RATE_LIMITED` with `Retry-After`;
    the resolver does NOT retry automatically (consistent with
    `investigation-outlook-cli.md §4.9`).
- **Effort: S-M** (the resolver is the bulk; paging helper is ~30 LoC).

#### A2. `$filter=DisplayName eq '<escaped>'` per level (REJECTED)

Same depth-first walk but with a server-side filter at each step. Rejected
because: (1) REST v2.0 is inconsistently deprecated and tenant-to-tenant
`$filter` support on `/childfolders` is spotty — the refined spec §12-A4
already acknowledges this and mandates client-side fallback; (2) ambiguity
detection (§6.4) still requires listing all matches, defeating the savings;
(3) single-quote escaping (`O''Brien`) must be implemented anyway. A1 gives
us the same capabilities with one code path instead of two.

#### A3. Fully-materialized tree cache (REJECTED)

"Download the whole folder tree once, resolve everything in memory." Rejected:
blows past the 50-page safety cap on big mailboxes, invalidates on every
concurrent Outlook-web action, has no place to live (session file is
auth-only), and the refined spec explicitly specifies serial, on-demand
resolution. It also makes `--first-match` + ambiguity rules harder to reason
about because the "current state" becomes a snapshot of varying freshness.

#### Well-known aliases

The v2.0 endpoint accepts the literal tokens `Inbox`, `SentItems`, `Drafts`,
`DeletedItems`, `Archive`, `JunkEmail`, `Outbox`, `MsgFolderRoot`, and
`RecoverableItemsDeletions` in the URL path as a substitute for the opaque
`Id`. This is already exercised by the existing `list-mail` command
(`src/commands/list-mail.ts:89` — `/api/v2.0/me/MailFolders/${encodeURIComponent(folder)}/messages`)
against `Inbox`, `SentItems`, `Drafts`, `DeletedItems`, `Archive` and returns
successfully. The resolver short-circuits at the root when the first segment
exactly matches an alias (no REST call needed to "resolve" Inbox — the URL is
emitted directly).

**Example request — list direct children of the root:**

```text
GET /api/v2.0/me/MailFolders?$top=100&$select=Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,TotalItemCount,WellKnownName
Authorization: Bearer <redacted>
Accept: application/json
```

**Example response (abridged):**

```json
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/MailFolders",
  "value": [
    {
      "Id": "AAMkAGI...Inbox",
      "DisplayName": "Inbox",
      "ParentFolderId": "AAMkAGI...root",
      "ChildFolderCount": 3,
      "UnreadItemCount": 12,
      "TotalItemCount": 402,
      "WellKnownName": "inbox"
    },
    {
      "Id": "AAMkAGI...Sent",
      "DisplayName": "Sent Items",
      "ParentFolderId": "AAMkAGI...root",
      "ChildFolderCount": 0,
      "UnreadItemCount": 0,
      "TotalItemCount": 58,
      "WellKnownName": "sentitems"
    }
  ],
  "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/MailFolders?$top=100&$skip=100&$select=..."
}
```

**Example — list children under Inbox:**

```text
GET /api/v2.0/me/MailFolders/AAMkAGI...Inbox/childfolders?$top=100&$select=Id,DisplayName,...
```

### B. Folder creation

#### B1. Lookup-then-create, idempotent on 409/400 (RECOMMENDED)

Flow for `create-folder Projects/Alpha --parent Inbox --create-parents --idempotent`:

1. Resolve parent anchor (`Inbox`) to an id (alias short-circuit, one call
   avoided).
2. Walk the path, segment by segment. For each segment:
   - `GET` children of the current parent, locate by DisplayName.
   - If found → advance (record `PreExisting: true`).
   - If not found and `--create-parents` is off AND we are not on the leaf
     → `UsageError FOLDER_MISSING_PARENT` (exit 2).
   - Otherwise: `POST /me/MailFolders/{parentId}/childfolders`,
     body `{"DisplayName": "<segment>"}`.
3. On a `POST` that returns 409 (or 400 with a duplicate-name error code — see
   §6.2 research topic): if `--idempotent` is set, re-list the parent's
   children to retrieve the id of the now-existing folder and proceed. If
   `--idempotent` is not set, raise the collision error (→ exit 6).

- **Pros**
  - Same HTTP shape as the rest of the tool (single `GET`/`POST` per step,
    each wrapped by `mapHttpError`).
  - `--create-parents` and `--idempotent` plug in cleanly as flags on the
    loop, not as server-side negotiation.
  - Ambiguity is impossible during create (we only ever create under a
    concrete parent id, and we just looked up that parent), so
    `FOLDER_AMBIGUOUS` cannot arise here.
- **Cons**
  - Race window: between the lookup and the `POST`, another Outlook client
    could create the same folder. The 409/400 branch handles this correctly
    when `--idempotent` is set; without `--idempotent`, the user sees an
    accurate collision error — which is the desired behaviour.
  - Two REST calls per segment (list-then-create) when the folder was absent
    — acceptable (same cost as any tree-mutation API in REST).
- **Failure modes**
  - The 400-vs-409 ambiguity (research topic §6.2). Without empirical pinning,
    the create-idempotent branch must inspect the upstream error body for a
    `"ErrorFolderExists"` or equivalent OData code and treat either status
    as collision.
  - A tenant that bans folder creation at the root (some Exchange Online
    policies do). Returns 403; we map to `UPSTREAM_HTTP_403` exit 5 without
    special-case.

**Example — create a top-level folder under the root:**

```text
POST /api/v2.0/me/MailFolders
Authorization: Bearer <redacted>
Content-Type: application/json
Accept: application/json

{"DisplayName": "Projects"}
```

**Example 201 response:**

```json
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/MailFolders/$entity",
  "Id": "AAMkAGI...Projects",
  "DisplayName": "Projects",
  "ParentFolderId": "AAMkAGI...root",
  "ChildFolderCount": 0,
  "UnreadItemCount": 0,
  "TotalItemCount": 0
}
```

**Example — create a nested folder:**

```text
POST /api/v2.0/me/MailFolders/AAMkAGI...Projects/childfolders
Content-Type: application/json

{"DisplayName": "Alpha"}
```

**Example collision (most tenants):**

```text
HTTP/1.1 409 Conflict
{"error":{"code":"ErrorFolderExists","message":"A folder with the specified name already exists."}}
```

Some tenants return 400 with the same `ErrorFolderExists` OData code. See §6.2.

#### B2. Create-then-lookup-on-fail (REJECTED)

"POST first, on error inspect and lookup." Rejected: symmetric cost to B1 in
the happy case (one POST vs. one GET + POST), but on the collision path it
requires parsing the upstream error body and then issuing a lookup — same
number of calls as B1 but with worse diagnostics (we did not intend to
create, yet the server log shows a POST). B1 is easier to reason about in
`--idempotent` mode.

#### B3. Server-side PATCH-upsert (REJECTED — not supported)

REST v2.0 does not offer an upsert primitive for `/MailFolders`. PATCH exists
only for already-known ids (rename) and is out of scope (NG1). Rejected.

### C. Move message

#### C1. `POST /me/messages/{id}/move` with `{"DestinationId": "<id-or-alias>"}` (RECOMMENDED)

The single wire shape for moving an email. The body accepts either a
well-known alias (`"DestinationId": "Archive"`) or an opaque folder id; we
confirm alias support is documented in the REST v2.0 reference (see reference
table below) and mirror what the web client does. The response is a **new**
message resource with a **new** `Id`.

- **Pros**
  - Single REST call per message — matches the serial-move constraint (NG4,
    NG10).
  - Destination alias support means we can let the resolver short-circuit
    for the common `move-mail --to Archive` case.
  - Mirrors what Outlook web does when a user drags a message between folders.
- **Cons**
  - The new-id semantics are a scripting footgun unless we surface the
    mapping explicitly. Refined spec §5.4 mandates the `{sourceId, newId}`
    pair per move — that's the required mitigation.
  - No native `$batch` support in v2.0 (refined NG4): N messages = N requests - N round-trips. The partial-failure shape (`moved[] / failed[] /
summary`) mirrors `download-attachments`.
- **Failure modes**
  - Source message already moved by another client → 404 on source id. Maps
    to `UPSTREAM_HTTP_404` exit 5 (or absorbed into `failed[]` under
    `--continue-on-error`).
  - Destination alias not recognized by the tenant (e.g. `Archive` not
    provisioned). Returns 400 with `ErrorInvalidIdMalformed` or similar.
    Maps to `UPSTREAM_HTTP_400` exit 5.
  - Destination id is valid but the message is system-locked (e.g. moving
    items out of `Outbox` that are currently sending) → 403. Maps to exit 5
    without retry.

**Example — move one message to a user folder:**

```text
POST /api/v2.0/me/messages/AAMkAGI...srcA/move
Authorization: Bearer <redacted>
Content-Type: application/json
Accept: application/json

{"DestinationId": "AAMkAGI...Alpha"}
```

**Example 201/200 response (new id highlighted):**

```json
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/Messages/$entity",
  "Id": "AAMkAGI...newA",
  "Subject": "Q1 budget review",
  "ReceivedDateTime": "2026-04-20T14:12:05Z",
  "ParentFolderId": "AAMkAGI...Alpha",
  "IsRead": true
}
```

**Example — move to a well-known alias:**

```text
POST /api/v2.0/me/messages/AAMkAGI...srcA/move
Content-Type: application/json

{"DestinationId": "Archive"}
```

(Alias support documented: see §6.1 research topic for status.)

#### C2. Client-side copy + delete (REJECTED)

"POST /copy to the destination, then DELETE the source." Rejected: two
round-trips per message, non-atomic (crash between calls leaves a duplicate),
and a copy's behaviour for some message types (drafts) differs subtly from
move. Outlook web uses `/move`; we should too.

#### C3. `$batch` multi-move in a single request (REJECTED — NG4)

Explicitly out of scope per refined spec NG4, and v2.0's `$batch` support is
partial and tenant-dependent. If future iterations add it, the partial-failure
shape stays the same.

### D. List messages in a folder

#### D1. Reuse the existing `list-mail` path, parameterized by folder id (RECOMMENDED)

The existing `list-mail.ts:89` builds
`/api/v2.0/me/MailFolders/${encodeURIComponent(folder)}/messages`. For the
well-known-alias fast path we keep it verbatim — `encodeURIComponent('Inbox')`
is `Inbox`, no change on the wire. For a folder id we do exactly the same
thing — Outlook accepts the opaque id in the same URL slot as the alias, the
refined spec §5.5 confirms this URL construction rule.

Flow:

1. If `--folder-id` is set → go straight to `/me/MailFolders/{id}/messages`
   (skip resolver).
2. Else if `--folder` matches an alias → existing fast path, unchanged.
3. Else → resolver to get an id, then path-build as in case 1.

- **Pros**
  - Zero wire-shape change for the existing test matrix; AC-LISTMAIL-WELLKNOWN-BACKCOMPAT
    is trivially satisfied.
  - The resolver is the only new moving part; the HTTP call is identical.
  - `$select`, `$orderby`, `$top`, all existing OData parameters still apply
    unchanged.
- **Cons**
  - None material. The only nuance is that the folder id must be URL-safe
    (it already is — Outlook opaque ids use only `[A-Za-z0-9=+/\-_.]`).
- **Failure modes**
  - A stale id → 404 on `/messages`. Maps to `UPSTREAM_HTTP_404` exit 5.
  - Caller passed both `--folder` and `--folder-id` → `UsageError BAD_USAGE`
    exit 2 (handled in command handler).

**Example — list messages by folder id:**

```text
GET /api/v2.0/me/MailFolders/AAMkAGI...Alpha/messages?$top=10&$orderby=ReceivedDateTime%20desc&$select=Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead,WebLink
```

(Response identical shape to the existing `list-mail` — `ODataListResponse<MessageSummary>`.)

#### D2. `GET /me/messages?$filter=ParentFolderId eq '...'` (REJECTED)

"Filter on the flat messages collection by parent folder id." Rejected: same
result, different code path, slower on the server side (tenants differ on
index coverage), and forces us to diverge from the well-known-alias fast
path. The existing URL shape is the one Outlook-web uses; we stay on it.

### E. Auth / session

The captured Bearer is an OWA web-app token with audience
`https://outlook.office.com`. Folder endpoints (`/MailFolders`,
`/MailFolders/{id}/childfolders`, `/messages/{id}/move`) share that audience
— they are part of the same REST v2.0 surface. **No new scopes are required.**
The existing `OutlookClient` 401-retry-once envelope applies verbatim: any
folder command's first 401 triggers exactly one `onReauthNeeded` run, and a
second 401 surfaces as `AUTH_401_AFTER_RETRY` exit 4. This is confirmed by
reading `src/http/outlook-client.ts:93-110` — the 401 logic is method-agnostic
(method name never appears in the branch), so extending `doGet` → `doRequest`
keeps the envelope unified.

No option rejected here — there is no alternative auth path for this
iteration.

### F. Rate limits & throttling

Outlook REST v2.0 uses the standard `429 Too Many Requests` + `Retry-After`
(seconds) header. The existing client already surfaces this as
`ApiError{code: 'RATE_LIMITED'}` with the `Retry-After` value in the message
(`src/http/outlook-client.ts:328-337`), which `mapHttpError` in `list-mail.ts`
re-maps to `UpstreamError{code: 'UPSTREAM_HTTP_429'}` exit 5.

For the recursive `list-folders` walk the concern is cumulative: N levels ×
M children per level × paging can issue dozens of requests. Two realistic
options:

#### F1. Surface 429 exit 5 immediately, no auto-retry (RECOMMENDED)

Mirrors how the existing `download-attachments` command handles 429 — user
re-runs after the `Retry-After` window. Pros: consistent with existing
behaviour; no hidden delays; scriptable (exit 5 + `Retry-After` in the error
body is enough for wrapper shells to sleep-and-retry). Cons: a long recursive
walk aborted at 80% completion wastes work — acceptable in v1.

#### F2. Automatic backoff + retry on 429 (REJECTED for this iteration)

Rejected: introduces hidden latency, requires a retry budget config, and
breaks symmetry with the rest of the tool. Can be added later as a shared
HTTP-layer feature for all commands at once.

---

## 3. Recommendation

**Primary (this iteration): the combination A1 + B1 + C1 + D1 + F1 with shared
paging and POST plumbing added to the existing `OutlookClient`.**

Justification (why this set, not the alternatives):

1. **Keeps a single REST shape per command.** Every folder command is one
   `GET` or one `POST` per step, wrapped by `mapHttpError`. Zero new fetch
   call sites. This is the hardest rule from the codebase scan (§Patterns to
   preserve #1) and the cheapest way to inherit the 401-retry envelope, the
   redact rules, the request-id capture, and the timeout handling.

2. **Isolates the one piece of genuine novelty — the resolver — in `src/folders/resolver.ts`.**
   The resolver owns path parsing, NFC + case-fold, well-known precedence,
   ambiguity detection, pagination awareness (via the client's `listAll`),
   and the 409/400-duplicate error classification. Every command
   (`list-folders`, `find-folder`, `create-folder`, `move-mail` destination
   resolution, `list-mail` non-alias folder) calls the resolver's
   `resolveFolder` or `createFolderPath`; no duplicate logic.

3. **Surfaces the new-id mapping for move explicitly.** The `moved[] /
failed[] / summary` shape is the direct mitigation for the "scripts chain
   old id to new id" footgun. It's also the exact blueprint already used by
   `download-attachments` (`saved[] / skipped[]`), so the formatter and
   failure-accumulation code patterns carry over.

4. **Requires no new mandatory config and no new exit codes.** Collisions
   route to the existing exit-6 slot (either via `IoError` re-use or a new
   `CollisionError` — refined §13 open question, but both land on exit 6).
   Ambiguity routes to the existing exit-2 slot via `UsageError`. Not-found
   routes to the existing exit-5 slot via `UpstreamError`. The taxonomy is
   preserved.

5. **Paging cap is implemented once, in the client, with an explicit code
   path.** The refined spec §7 50-page cap becomes one method on
   `OutlookClient` (`listAll<T>(path, query?)` or `getPaged`), raising
   `UpstreamError{code: 'UPSTREAM_PAGINATION_LIMIT'}`. Every folder
   enumeration inherits this; `list-folders --recursive` is bounded
   automatically.

Risks the implementer must actively handle:

1. **Race window during create-parents + idempotent.** Two concurrent
   `create-folder A/B/C --create-parents --idempotent` runs can both succeed
   (lookup misses, POST wins or collides). Every collision path must re-list
   to recover the id; do not assume the POST's response is authoritative when
   the status is 409/400.

2. **DestinationId alias support is underdocumented.** The research target
   §6.1 asks us to pin down whether the body value `"Archive"` (rather than
   an opaque id) is accepted by the `/move` endpoint on all target tenants.
   The safe implementation: default to resolving every alias to an id
   client-side, and opt in to alias pass-through only for a known list.

3. **409 vs 400 on duplicate create.** See §6.2. Implementation must inspect
   the upstream OData `code` field (`ErrorFolderExists` or similar) on both
   statuses, not just the HTTP status number, before concluding "collision
   vs real bad request."

4. **Recursive `list-folders` spanning a huge tree.** The 50-page cap (refined
   §7) is a per-folder-children bound, not a whole-tree bound. A pathological
   tree could still fan out hundreds of `GET`s. The command must warn in
   stderr after, say, 500 total requests, and surface a fatal
   `UPSTREAM_PAGINATION_LIMIT` if a _single_ level exceeds 50 pages. Keep
   the tree fan-out visible.

5. **Secrets in error bodies.** `POST /MailFolders` / `POST /move` echo request
   bodies into some error responses. The existing `truncateAndRedactBody` in
   `src/http/errors.ts:151` must apply unchanged — no folder-specific error
   path can bypass it.

**Secondary (future iteration): Microsoft Graph migration.**

The refined spec §12-A1 explicitly keeps Graph out of scope for this
iteration. If that constraint is ever lifted, the equivalent Graph endpoints
(`/me/mailFolders`, `/me/mailFolders/{id}/childFolders`, `/me/messages/{id}/move`)
have identical semantics and slightly better OData filter coverage. The
folder resolver would port over with a base-URL change and PascalCase →
camelCase field renames only. Track as a future enhancement.

---

## 4. Implementation strategy (for the recommended approach)

### 4.1 HTTP client extension — `post` + `listAll`

**Recommendation:** in `src/http/outlook-client.ts`, refactor the private
`doGet` into a generic `doRequest(method, path, body?, query?)`, add a public
`post<TBody, TRes>(path, body, query?)` method to the `OutlookClient`
interface, and add a `listAll<T>(path, query?)` method that follows
`@odata.nextLink` up to 50 pages.

Reasoning:

- The 401-retry-once envelope lives in `doGet` lines 93-110 today. Hoisting
  it to `doRequest` and parameterizing on method + optional body is ~30 LoC
  and the only safe way to share re-auth across GET and POST.
- `buildUrl`, `buildHeaders`, `executeFetch`, `handleSuccessOrThrow`,
  `throwForResponse`, `mapFetchException` — **unchanged**. They are all
  method-agnostic.
- `post` must set `Content-Type: application/json` and serialize the body as
  `JSON.stringify(body)`. The existing `buildHeaders` returns a fresh object
  per request, so adding `Content-Type` only on POST is safe.
- `listAll<T>(path, query?)` issues the first `GET` with `query`, then for
  each subsequent page it follows the _full_ URL in `@odata.nextLink` (which
  already contains `$skip` / `$skiptoken`, so we do NOT re-apply `query`). A
  page counter gates the 50-cap; on overflow it throws
  `ApiError{code: 'PAGINATION_LIMIT'}` which `mapHttpError` maps to
  `UpstreamError{code: 'UPSTREAM_PAGINATION_LIMIT'}` — extending the
  existing `code` vocabulary rather than introducing a new class.
- `@odata.nextLink` is absolute; Outlook emits it on the same host
  (`outlook.office.com`). The helper must reject links that point off-host
  (defense in depth) before calling `executeFetch`.

### 4.2 Folder resolver — `src/folders/resolver.ts`

**Recommendation:** one module, five exported functions:

- `parseFolderPath(input): string[]` — split on `/`, unescape `\/` → `/` and
  `\\` → `\`; enforce the 16-segment cap and the "no empty segment" rule;
  NFC-normalize every segment; throw `UsageError FOLDER_PATH_INVALID` on
  violations.
- `buildFolderPath(segments): string` — inverse, used for the materialized
  `Path` field in `list-folders --recursive` output.
- `matchesWellKnownAlias(input): string | null` — single-pass exact match
  against the §6.2 alias list; returns the canonical form (`Inbox`,
  `SentItems`, ...) or null.
- `listChildren(client, parentId, { top, includeHidden }): Promise<FolderSummary[]>`
  — wraps `client.listAll(...)`; materialized array of direct children.
- `resolveFolder(client, spec, { caseSensitive, includeHidden, firstMatch }): Promise<ResolvedFolder>`
  — the path-walk workhorse. Returns the final folder plus `resolvedVia`
  discriminator (`'wellknown' | 'path' | 'id'`).
- `createFolderPath(client, { anchorId, segments, createParents, idempotent }): Promise<CreateResult>`
  — the create loop described in B1.

Reasoning:

- All path / alias / NFC / ambiguity logic lives here. Commands never
  recompute any of it.
- `resolveFolder` takes the same `OutlookClient` that commands already own;
  no extra wiring.
- Every resolver path that detects "not found" throws `UpstreamError{code:
'UPSTREAM_FOLDER_NOT_FOUND'}`. Every path that detects "ambiguity" throws
  `UsageError{code: 'FOLDER_AMBIGUOUS'}` (unless `firstMatch` is set). Both
  mapping rules are refined spec §10.

### 4.3 Collision error class

**Recommendation:** introduce a new `CollisionError extends OutlookCliError`
in `src/config/errors.ts` with `exitCode = 6` and `code = 'FOLDER_ALREADY_EXISTS'`.

Reasoning:

- The refined spec §13 open question asks whether to route folder collisions
  through `IoError` or a new class. A new class is cleaner because the cause
  is _not_ filesystem IO (the existing exit-6 path) and because new callers
  (`cli.ts formatErrorJson` / `exitCodeFor`) gain a discriminated
  `instanceof CollisionError` branch with no ambiguity.
- One-line addition to `exitCodeFor` (`if (err instanceof CollisionError)
return 6`) and `formatErrorJson` (serialize `{code, path?, ...}`).
- If the plan phase decides against a new class, the fallback is `IoError`
  with a folder-specific code. Both preserve exit 6.

### 4.4 Command shape — one per subcommand

Every new command file (`list-folders.ts`, `find-folder.ts`,
`create-folder.ts`, `move-mail.ts`) follows the canonical shape documented
in `codebase-scan-folders.md §2.1`:

```text
1. validate opts against CliConfig defaults
2. ensureSession(deps) → SessionFile
3. deps.createClient(session) → OutlookClient
4. build resolver inputs; call resolver
5. build REST call(s); wrap in try { ... } catch (err) { throw mapHttpError(err); }
6. return typed value
```

`move-mail` is the only command that loops (the multi-id case). It uses the
same accumulator shape as `download-attachments`: `moved[]` / `failed[]` /
`summary` plus a discriminator on whether `--continue-on-error` absorbs
individual failures.

### 4.5 `list-mail` extension — preserve the well-known fast path

**Recommendation:** add `--folder-id` and `--folder-parent` to
`ListMailOptions`. Widen the `--folder` acceptor so:

- If value is in `ALLOWED_FOLDERS` → existing fast path, no resolver, no
  behaviour change (satisfies AC-LISTMAIL-WELLKNOWN-BACKCOMPAT).
- Else if value matches any other well-known alias from §6.2 → fast path
  against that alias.
- Else → resolve via the resolver (path walk), then hit
  `/MailFolders/{id}/messages`.

Reasoning:

- The existing `ALLOWED_FOLDERS` constant (`src/commands/list-mail.ts:37`)
  stays; it becomes the "fast-path list" rather than the "allowlist." The
  validation check at line 73 is widened but still rejects empty / invalid
  strings.
- `--folder-id` and `--folder` are mutually exclusive; validation raises
  `UsageError BAD_USAGE` (exit 2).

### 4.6 Output formatting

**Recommendation:** follow the existing `ColumnSpec` pattern; add three
constants in `cli.ts` next to `LIST_MAIL_COLUMNS`:

- `LIST_FOLDERS_COLUMNS` — `Path | Unread | Total | Children | Id` (no
  `maxWidth` on `Id`).
- `CREATE_FOLDER_COLUMNS` — applied to `result.created`: `Path | Id |
PreExisting`.
- `MOVE_MAIL_COLUMNS` — `Source Id | New Id | Status | Error`.

`find-folder` returns a single object; the current `emitResult` in `cli.ts`
falls back to JSON for non-array payloads. Acceptable in v1. A future
`--kv-output` mode is a small extension to the formatter if the refined
spec's "two-line key/value" table for `find-folder` becomes critical.

### 4.7 Error taxonomy extensions

Additions per refined spec §10 (no new exit codes, new codes on existing
classes):

| Class                  | New code                    | Exit |
| ---------------------- | --------------------------- | ---- |
| `UsageError`           | `FOLDER_AMBIGUOUS`          | 2    |
| `UsageError`           | `FOLDER_MISSING_PARENT`     | 2    |
| `UsageError`           | `FOLDER_PATH_INVALID`       | 2    |
| `UpstreamError`        | `UPSTREAM_FOLDER_NOT_FOUND` | 5    |
| `UpstreamError`        | `UPSTREAM_FOLDER_AMBIGUOUS` | 5    |
| `UpstreamError`        | `UPSTREAM_PAGINATION_LIMIT` | 5    |
| `CollisionError` (new) | `FOLDER_ALREADY_EXISTS`     | 6    |

`AuthError` and `IoError` unchanged.

### 4.8 Logging & redaction

The `X-AnchorMailbox` header continues to carry `PUID:<puid>@<tenantId>` —
which is not a secret but is PII. It is already emitted as-is. The Bearer and
cookie values remain off-limits. The existing `redactString` (from
`src/util/redact.ts`) already handles the only two leakage vectors (the
Authorization header and cookie values) in `buildHeaders`; no folder-specific
redact rules are needed.

### 4.9 Error taxonomy wire-up in `cli.ts`

Two edits:

1. `formatErrorJson` (`src/cli.ts:297`) — add an `if (err instanceof
CollisionError)` branch that serializes `{code, path?, parentId?}`.
2. `exitCodeFor` (`src/cli.ts:359`) — add `if (err instanceof
CollisionError) return 6`.

Both are single-line additions; the existing `OutlookCliError` fallback
already covers `exitCode` for unknown subclasses but explicit branches make
the JSON shape deterministic.

---

## 5. Risk register

| Risk                                                                         | Likelihood | Impact | Mitigation                                                                                                                                                    |
| ---------------------------------------------------------------------------- | ---------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Outlook returns **400** (not 409) on duplicate-name create                   | Medium     | Medium | Inspect upstream OData `code` string (`ErrorFolderExists`) on both 400 and 409; map both to `FOLDER_ALREADY_EXISTS` when `--idempotent`. Research topic §6.2. |
| `/move` `DestinationId` rejects a well-known alias on some tenant            | Low-Medium | Low    | Default: resolve alias → id client-side before calling `/move`. Only use alias pass-through when an explicit `--raw-alias` flag is set (not part of v1).      |
| `@odata.nextLink` points off-host after a redirect                           | Very Low   | High   | Reject `nextLink` that is not `https://outlook.office.com/...`; raise `UPSTREAM_PAGINATION_LIMIT` with explanatory message.                                   |
| Recursive walk hits 429 mid-tree                                             | Medium     | Low    | Existing `mapHttpError` surfaces `UPSTREAM_HTTP_429` with `Retry-After`. No auto-retry in v1 (consistent with rest of tool). User re-runs.                    |
| Folder DisplayName contains a raw `/` that the user forgot to escape         | Medium     | Medium | Resolver raises `UPSTREAM_FOLDER_NOT_FOUND` on the wrong segment; help text on `find-folder` lists the escape rules. AC-PATH-ESCAPE covers the happy path.    |
| Ambiguous path matches both a user folder and a well-known alias at the root | Low        | Medium | Spec §6.2 pins "well-known wins at root"; user must pass `--parent MsgFolderRoot --first-match` to reach a shadowed user folder.                              |
| `POST /move` succeeds but the response body is empty (tenant variant)        | Very Low   | Medium | Treat an empty/absent `Id` in the response as `UPSTREAM_HTTP_<status>` with message "move response missing new id"; scripts see `failed[]` not `moved[]`.     |
| Pagination cap hit on a legitimate huge tree                                 | Low        | Low    | `UPSTREAM_PAGINATION_LIMIT` exit 5 with an actionable message ("use `--parent <sub-folder>` to walk a smaller subtree"). Documented in help text.             |
| Concurrent `create-folder` runs collide on the same leaf                     | Very Low   | Low    | 409/400 branch re-lists and returns the existing id under `--idempotent`; without `--idempotent`, the user sees an accurate collision error (desired).        |
| Bearer token expires mid-recursive-walk                                      | Low        | Medium | Existing 401-retry-once envelope handles it transparently; the current call is retried after re-auth, subsequent pages use the new token (mutable session).   |
| Hidden-by-default folders leak when `$select` lacks `IsHidden`               | Low        | Low    | Always include `IsHidden` in the default `$select`; resolver filters at the materialization boundary unless `--include-hidden` is set.                        |
| `--first-match` silently hides the "real" target                             | Low        | Medium | Flag is documented as a foot-gun; tests include `AC-FOLDER-AMBIGUOUS` without `--first-match` to ensure exit 2 is the default.                                |

---

## 6. Technical Research Guidance

Research needed: **Yes**. Three items are NOT pinned down well enough to commit to code paths. Each is
bounded and should take at most a short, empirical probe against a live
tenant plus cross-checking Microsoft docs.

### Topic 1: `/move` `DestinationId` — alias acceptance on Outlook REST v2.0

- **Why**: Refined spec §5.4 allows `--to Archive` to pass straight through
  without client-side resolve. The current HTTP code path has never exercised
  alias-in-POST-body; the existing alias use has been URL-in-path only
  (`/MailFolders/Inbox/messages`). A tenant that rejects alias-in-body forces
  us to always pre-resolve aliases to ids before calling `/move`.
- **Focus**:
  - Live probe: `POST /api/v2.0/me/messages/{id}/move {"DestinationId":"Archive"}`
    against the existing captured session. Observe 200 vs 400.
  - Confirm the full alias list that is accepted in `DestinationId` (v2.0
    REST reference historically lists `Inbox`, `Drafts`, `SentItems`,
    `DeletedItems`; `Archive` / `JunkEmail` / `Outbox` may or may not be
    accepted — tenant-dependent).
  - Microsoft Graph equivalent (`POST /me/messages/{id}/move
{"destinationId":"archive"}` lowercased on Graph) for cross-reference.
- **Depth**: Overview (one empirical call + one doc cross-ref is enough).
- **Relevance**: Determines whether §4.2's `resolveFolder` is unconditionally
  called for every `--to` value, or short-circuits when the value is an alias.

### Topic 2: Duplicate-folder create — 409 vs 400 on `POST /childfolders`

- **Why**: Refined spec §7 explicitly flags "some tenants return 400" for
  the collision path. The implementer cannot rely on HTTP status alone; the
  OData `code` field in the error body must be inspected. The shape of that
  error body on v2.0 (not Graph) is not pinned down in public docs.
- **Focus**:
  - Live probe: `POST /api/v2.0/me/MailFolders` with a DisplayName that
    already exists. Capture the exact response body on 409 (dev account, probably) and
    on 400 (if any tenant reproduces).
  - Confirm the OData `code` string — `ErrorFolderExists`, `ErrorCreateItemAccessDenied`,
    or a newer code — that identifies a duplicate-name collision.
  - Check whether the error body shape is `{ "error": { "code": "...", "message": "..." } }`
    consistently (Graph and v2.0 are almost identical here, but "almost" is
    what bites).
- **Depth**: Overview (one or two probes + the error-response schema).
- **Relevance**: Determines the exact condition under which `--idempotent`
  swallows the error. Without this pinned, the implementation must pattern-
  match on a broader set of codes, which risks false positives (swallowing
  a legitimate 400 for a different reason).

### Topic 3: Recursive `/childfolders` pagination — `$top` / `$skip` vs `@odata.nextLink`

- **Why**: The existing `OutlookClient` has never paged. The refined spec §7
  puts a 50-page cap but does not specify whether `@odata.nextLink` on v2.0
  uses `$skip`, `$skiptoken`, or a combination; whether the cursor is stable
  under concurrent mailbox mutation; and whether Outlook honors a caller-
  supplied `$skip` when also returning a `nextLink` (some tenants return
  truncated pages with `$top` ignored). Pinning this down before writing the
  `listAll<T>` helper is cheap and avoids a class of subtle bugs.
- **Focus**:
  - Live probe: `GET /api/v2.0/me/MailFolders?$top=2` against an account with
    > 2 top-level folders. Inspect `@odata.nextLink` — is it `...?$top=2&$skip=2`?
    > `...?$skiptoken=...`? Something else?
  - Follow the link verbatim (no query rewriting) and confirm the next page
    advances correctly.
  - Test with `$top=1` (boundary) and with `$top=100` (request page larger
    than any likely collection — confirm no `nextLink` is emitted).
  - Confirm that the `@odata.nextLink` is on the same host
    (`outlook.office.com`).
  - Cross-reference: Graph's paging is documented; v2.0's paging should be
    identical but the MS docs for v2.0 are sparse.
- **Depth**: Overview-to-intermediate (three empirical calls + one doc
  cross-ref).
- **Relevance**: Determines the exact loop shape of `listAll<T>`. If
  `@odata.nextLink` is always self-contained (the common case), the helper
  is trivial: `while (url) { page = GET url; yield page.value; url =
page['@odata.nextLink'] }`. If the tenant returns `$skiptoken`-based
  cursors that require the caller to remember `$top`, the helper needs to
  preserve the original query bag. This topic is the single gating item for
  `list-folders --recursive`.

---

## Summary

- **Primary approach: A1 + B1 + C1 + D1 + F1** — client-side path walk for
  lookup, lookup-then-create for idempotent creation, `POST /move` with
  explicit new-id surfacing, reuse the existing `/MailFolders/{id}/messages`
  URL for in-folder listing, surface 429 as exit 5 without auto-retry.
- **Shared plumbing additions**: `OutlookClient.post`, `OutlookClient.listAll`
  with the 50-page cap, a new `CollisionError` class at exit 6, and a new
  `src/folders/resolver.ts` module that owns every piece of path / alias
  semantics.
- **Existing machinery preserved**: `ensureSession`, `mapHttpError`,
  `UsageError`, `buildUrl`, `buildHeaders`, `executeFetch`,
  `handleSuccessOrThrow`, `throwForResponse`, `redactString`, the 401-retry-
  once envelope — unchanged.
- **Open research (three topics, all short and empirical)**: `/move`
  `DestinationId` alias acceptance, 409-vs-400 for duplicate folder create,
  and `@odata.nextLink` paging shape on v2.0 `/MailFolders`. All flagged in
  §6 for Phase 3b before implementation begins.

## References

| #   | Source                                                  | URL                                                                                                                       | What was learned                                                                                                                                                                                                                           |
| --- | ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1   | Microsoft Docs — Outlook Mail REST v2.0 — folders       | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations            | Endpoint paths for `/me/MailFolders`, `/me/MailFolders/{id}/childfolders`, `POST` body shape `{"DisplayName": "..."}`, well-known alias tokens accepted in the URL path.                                                                   |
| 2   | Microsoft Docs — Outlook REST v2 — Move message         | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/message-rest-operations         | `POST /me/messages/{id}/move` body shape `{"DestinationId": "..."}`, response returns the moved message with a NEW `Id`, alias acceptance in `DestinationId` (historical; research topic §6.1 needed to confirm current tenant behaviour). |
| 3   | Microsoft Graph — Create MailFolder (schema-equivalent) | https://learn.microsoft.com/en-us/graph/api/user-post-mailfolders?view=graph-rest-1.0                                     | Collision error code `ErrorFolderExists` on duplicate-name create; error body shape `{ "error": { "code": "...", "message": "..." } }`. Graph equivalent cross-referenced because v2.0 docs are sparse on the error body.                  |
| 4   | Microsoft Graph — Move message (schema-equivalent)      | https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0                                              | Confirms the "new id after move" semantics on Graph; v2.0 behaviour is identical. Used to corroborate refined spec §5.4 `moved[{sourceId,newId}]` requirement.                                                                             |
| 5   | Microsoft Graph — MailFolder list children              | https://learn.microsoft.com/en-us/graph/api/mailfolder-list-childfolders?view=graph-rest-1.0                              | `$top` / `@odata.nextLink` paging shape on `/childFolders`. Pinning v2.0 parity is research topic §6.3.                                                                                                                                    |
| 6   | OData v4 Specification — nextLink                       | https://docs.oasis-open.org/odata/odata/v4.01/os/part1-protocol/odata-v4.01-os-part1-protocol.html#sec_ServerDrivenPaging | `@odata.nextLink` is an absolute URL that the client must follow verbatim; generic OData behaviour used to anchor the expected paging semantics.                                                                                           |
| 7   | `codebase-scan-folders.md` (this project)               | `docs/reference/codebase-scan-folders.md`                                                                                 | Exact plug-in points: `src/cli.ts:486-507` (list-mail registration), `src/commands/list-mail.ts:37-104` (resolver integration site), `src/http/outlook-client.ts:80-117` (401 envelope + generalize to `doRequest`).                       |
| 8   | `investigation-outlook-cli.md` (this project)           | `docs/design/investigation-outlook-cli.md`                                                                                | Calibration for tone, structure, risk-register shape, technical-research-guidance shape, error-code taxonomy used here.                                                                                                                    |

## Original Request

> I want you to add support to search and create folders, move emails to folders, list emails in folders

Refined at: `docs/design/refined-request-folders.md` — full scope, subcommand
shapes, CLI surface, error-mapping rules, and acceptance criteria.

Absolute output path: `<upstream-repo>/docs/design/investigation-folders.md`
