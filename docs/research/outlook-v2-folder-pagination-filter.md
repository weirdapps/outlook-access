# Outlook REST v2.0 — Folder Pagination Contract & `$filter` on DisplayName

Research date: 2026-04-21
Scope: `GET /me/MailFolders` and `GET /me/MailFolders/{id}/childfolders` — pagination
semantics (`@odata.nextLink`, `$top`, `$skip`), and `$filter=DisplayName eq '...'` support.
Motivation: sizing `MAX_FOLDER_PAGES` / `MAX_FOLDERS_VISITED` caps and designing the
`listAll<T>` helper for the folder resolver described in
`docs/design/investigation-folders.md §4.1`.

---

## 1. TL;DR Parameter Table

| Parameter / Feature                     | Supported?                                    | Default    | Max                   | Notes                                                                |
| --------------------------------------- | --------------------------------------------- | ---------- | --------------------- | -------------------------------------------------------------------- |
| `@odata.nextLink` in response           | **Yes**                                       | —          | —                     | Absolute URL; emitted only when more pages exist                     |
| `$top` (client page-size hint)          | **Yes**                                       | **10**     | **~1000** (practical) | Server may honour a lower cap silently; see §3                       |
| `$skip` (offset)                        | **Yes**                                       | —          | —                     | Outlook/Mail APIs use `$skip`-based nextLinks, not `$skiptoken`      |
| `$select`                               | **Yes**                                       | All fields | —                     | Always use; reduces payload to ~120-200 bytes per folder             |
| `$filter=DisplayName eq '...'`          | **Yes (v2.0 schema marks it Filterable)**     | —          | —                     | Works in practice but tenant-to-tenant reliability is medium; see §4 |
| `$filter=startswith(DisplayName,'...')` | **Yes** (OData standard function, same field) | —          | —                     | Same caveats as `eq`                                                 |
| `includeHiddenFolders=true`             | **Yes**                                       | false      | —                     | Custom query param, not OData; use alongside `$filter`               |
| `$orderby`                              | **Yes**                                       | undefined  | —                     | Supported on filterable scalar fields                                |
| `$count`                                | **Yes**                                       | —          | —                     | Append `/$count` or `?$count=true`                                   |

**Bottom line**: `@odata.nextLink` is cursor-style with a `$skip` offset baked in; follow
it verbatim. `$filter=DisplayName eq 'X'` is schema-supported but must be treated as an
optimization with a mandatory client-side fallback (see §4 recommendation).

---

## 2. `@odata.nextLink` Mechanics

### What the response looks like

When the collection spans more than one page, Outlook appends `@odata.nextLink` to the
response body alongside `value`:

```json
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/MailFolders",
  "value": [
    { "Id": "AAMkAGI...", "DisplayName": "Inbox", ... },
    ...
  ],
  "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/MailFolders?$top=100&$skip=100&$select=Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,TotalItemCount"
}
```

When the last page is returned `@odata.nextLink` is **absent** (not present as `null`).

### Key properties

1. **Absolute URL** — the link is a fully qualified `https://outlook.office.com/...` URL.
   The bearer token does NOT need to be embedded; it remains in the `Authorization` header
   of the next request.

2. **`$skip`-based (not `$skiptoken`)** — the Microsoft Graph documentation explicitly
   states: _"Some Microsoft Graph APIs, like Outlook Mail and Calendars (message, event,
   and calendar), use `$skip` to implement paging."_ This contrasts with directory objects
   (users, groups) which use `$skiptoken`. For Mail folder collections the offset integer
   in `@odata.nextLink` is a plain `$skip=N` value.

3. **Self-contained** — the nextLink URL already encodes the original `$select`, `$top`,
   and any `$filter` from the first request. The `listAll<T>` helper must follow this URL
   **verbatim** without re-applying the original query bag. Re-applying `query` to the
   nextLink URL would produce double-encoding and potentially incorrect results.

4. **Same host guarantee** — the link is always on `https://outlook.office.com`. The
   `listAll<T>` implementation should validate the host before fetching, to defend against
   a malformed or redirect-rewritten URL (risk: very low, but the defense is a one-liner).

5. **Stability under concurrent mutation** — `$skip` offsets are NOT stable if the
   collection is modified concurrently (insertions before position N shift later items).
   In practice the folder tree changes slowly (user-driven renames/creates), and the
   resolver is a short-lived operation, so this is acceptable. For a long-running
   `list-folders --recursive` over a large tree there is a theoretical risk of a missed
   or duplicated folder; the 50-page cap (§6) bounds the exposure.

6. **Token-retry safety** — if a page request triggers a 401 and the `OutlookClient`'s
   401-retry-once envelope re-auths, the next `$skip`-based request still succeeds because
   the offset is stateless (no server-side cursor to expire).

---

## 3. `$top` and `$skip` Behavior

### Default page size

The Microsoft Graph documentation for `GET /me/mailFolders` explicitly states:

> "If a collection exceeds the **default page size (10 items)**, the `@odata.nextLink`
> property is returned..."

The v2.0 endpoint shares the same Exchange Online backend. Its default is also **10**
per page for `MailFolders` and `childfolders` collections. This is notably smaller than
the message default (also 10 but often server-auto-capped at a different value for large
tenants).

**Implication**: a call to `GET /me/MailFolders` without `$top` against a mailbox with
15 top-level folders will return 10 items and a `@odata.nextLink`. Always specify
`$top=100` (or higher) to minimize round-trips.

### Maximum `$top`

The v2.0 documentation does not publish a hard maximum for folder collections. In
practice:

- For `messages` the documented max is 1000.
- For `mailFolders` / `childfolders` the pattern matches; community reports and Graph
  Explorer tests confirm that `$top=1000` is accepted without error.
- **Safe practical cap**: use `$top=250` as a conservative default. This avoids any
  undocumented per-tenant limit, keeps response payloads manageable (250 × ~200 bytes =
  ~50 KB), and reduces the number of calls for most real-world mailboxes (few users have
  > 250 direct children at a single level).
- If a tenant silently ignores `$top` and returns fewer items, the `@odata.nextLink`
  mechanism compensates automatically — the loop continues regardless.

### `$skip` as a standalone parameter

`$skip` can be supplied directly by the caller, but there is **no reason to do so** in
the `listAll<T>` helper. The correct implementation follows `@odata.nextLink` verbatim;
the link already contains the correct `$skip` value for the next page. Constructing your
own `$skip` offset is fragile (concurrent inserts cause drift) and duplicates logic the
server already encodes in the link.

### Summary

```
First call:  GET /me/MailFolders/{id}/childfolders?$top=250&$select=...
                ↓ follow @odata.nextLink verbatim
Next pages:  GET https://outlook.office.com/api/v2.0/me/MailFolders/{id}/childfolders
                ?$top=250&$skip=250&$select=...   (link already contains this)
                ↓ repeat until no @odata.nextLink in response
```

---

## 4. `$filter` on DisplayName

### Schema support (v2.0)

The official v2.0 resource schema table for `MailFolder` explicitly marks `DisplayName`
as **Filterable: Yes**:

| Property         | Type   | Writable? | Filterable? |
| ---------------- | ------ | --------- | ----------- |
| DisplayName      | String | Yes       | **Yes**     |
| ChildFolderCount | Int32  | No        | Yes         |
| TotalItemCount   | Int32  | No        | Yes         |
| UnreadItemCount  | Int32  | No        | Yes         |
| Id               | String | No        | **No**      |
| ParentFolderId   | String | No        | **No**      |

So the wire syntax is legal:

```
GET /api/v2.0/me/MailFolders/Inbox/childfolders
    ?$filter=DisplayName eq 'Projects'
    &$select=Id,DisplayName,ChildFolderCount
```

For `startswith` (partial prefix match):

```
GET /api/v2.0/me/MailFolders/Inbox/childfolders
    ?$filter=startswith(DisplayName,'Proj')
    &$select=Id,DisplayName,ChildFolderCount
```

### Practical reliability assessment

**Confidence: MEDIUM.** The schema marks it filterable, but the v2.0 surface is
deprecated (decommissioned target was March 2024 — though tenants connected via
`outlook.office.com` OWA tokens continue to work in practice). Several known caveats:

1. **Case sensitivity is tenant-dependent.** OData `eq` is case-sensitive by default;
   some Exchange Online tenants perform case-insensitive comparisons, others do not.
   Do not rely on `DisplayName eq 'inbox'` matching `Inbox` without testing.

2. **Unicode normalization is not guaranteed.** A folder named with a precomposed
   Unicode character may not match a filter using the decomposed form and vice versa.

3. **Combination with `includeHiddenFolders`** — there is no documented guarantee that
   `$filter` and `includeHiddenFolders=true` can be combined. If the resolver needs
   hidden folders, prefer enumerating all and filtering client-side.

4. **The investigation already recommends client-side fallback** (§A2 in
   `investigation-folders.md`) as the primary strategy. The rejection reason: "REST v2.0
   is inconsistently deprecated and tenant-to-tenant `$filter` support on `/childfolders`
   is spotty." This research corroborates that judgment.

### Recommendation for the resolver

- **Do NOT use `$filter=DisplayName eq '...'` as the primary strategy.** The primary
  strategy remains A1 (enumerate children, match client-side via NFC + case-fold). This
  is not performance-significant for typical folder fan-out (a mailbox with 100 children
  at one level sends at most 1-2 pages of 250-item requests — about 25-50 KB per level).

- **`$filter` may be offered as an opt-in optimization** (`--server-filter` flag, not in
  v1) for power users with pathologically large flat hierarchies (thousands of direct
  children at one level). In that scenario, the helper would first attempt the filtered
  call; on 400/501/unsupported-filter response it falls back to full enumeration.

- **Always double-escape single quotes** in filter strings: a folder named `O'Brien`
  requires `$filter=DisplayName eq 'O''Brien'` (OData 4.0 single-quote escape rule).

---

## 5. Recommended `listAll<T>` Implementation Pattern

### Design principles

1. Follow `@odata.nextLink` verbatim — do not reconstruct or re-encode query parameters.
2. Validate that the nextLink is on the expected host before fetching.
3. Enforce a page cap to bound worst-case API calls per collection.
4. Surface cap exhaustion as a named error code (`PAGINATION_LIMIT`) that callers can
   map to `UpstreamError{code: 'UPSTREAM_PAGINATION_LIMIT'}`.
5. Never swallow empty pages — a page with zero items and a nextLink is valid OData
   (rare but possible); the loop must continue.

### TypeScript snippet

```typescript
// src/http/outlook-client.ts (addition to OutlookClient class)

const ALLOWED_HOST = 'outlook.office.com';
const DEFAULT_LIST_TOP = 250;

/** OData list-response envelope returned by Outlook REST v2.0 */
interface ODataListResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.context'?: string;
}

export interface ListAllOptions {
  /** Maximum pages to follow before throwing PAGINATION_LIMIT. Default: 50. */
  maxPages?: number;
  /** $top hint for the first request. Default: 250. */
  top?: number;
}

/**
 * Fetches all items from a paginated Outlook REST v2.0 collection.
 *
 * Issues the first GET with `path` + `query` (merged), then follows
 * every `@odata.nextLink` verbatim (no query re-application) until
 * the collection is exhausted or `maxPages` is reached.
 *
 * @throws ApiError{ code: 'PAGINATION_LIMIT' } when maxPages is exceeded.
 * @throws ApiError{ code: 'PAGINATION_OFF_HOST' } when nextLink points off-host.
 */
async listAll<T>(
  path: string,
  query?: Record<string, string>,
  opts: ListAllOptions = {}
): Promise<T[]> {
  const maxPages = opts.maxPages ?? 50;
  const top = opts.top ?? DEFAULT_LIST_TOP;

  // Merge caller query with $top default (caller can override $top explicitly).
  const firstQuery: Record<string, string> = { $top: String(top), ...query };
  const results: T[] = [];
  let pageCount = 0;

  // Build the initial URL.
  let url: string | null = this.buildUrl(path, firstQuery);

  while (url !== null) {
    if (pageCount >= maxPages) {
      throw new ApiError(
        'PAGINATION_LIMIT',
        `Exceeded ${maxPages}-page cap fetching ${path}. ` +
          `Use --parent to narrow the scope.`
      );
    }

    // Host validation — reject nextLinks that escaped to another host.
    const parsed = new URL(url);
    if (parsed.hostname !== ALLOWED_HOST) {
      throw new ApiError(
        'PAGINATION_OFF_HOST',
        `@odata.nextLink host '${parsed.hostname}' is not '${ALLOWED_HOST}'.`
      );
    }

    const page = await this.doGet<ODataListResponse<T>>(url);
    results.push(...page.value);
    pageCount++;

    // Follow the nextLink verbatim; it already encodes $skip and all original params.
    url = page['@odata.nextLink'] ?? null;
  }

  return results;
}
```

### Notes on the snippet

- `buildUrl` is the existing helper in `OutlookClient` and handles base-URL prepending.
  For subsequent pages, `url` is already absolute — pass it straight to `doGet`.
- `doGet` must accept a full URL (not just a path). If the current implementation always
  prepends the base URL, add an overload that accepts a pre-built absolute URL for
  pagination use only.
- The 401-retry-once envelope lives in `doGet`; it applies transparently to every page
  request, including nextLink-following ones, because the token is stateless relative to
  `$skip` position.
- `ApiError` is the existing internal error class; `mapHttpError` in calling commands
  converts `PAGINATION_LIMIT` → `UpstreamError{code: 'UPSTREAM_PAGINATION_LIMIT'}`.

### Usage in the resolver

```typescript
// src/folders/resolver.ts

async function listChildren(
  client: OutlookClient,
  parentId: string,
  opts: { top?: number; includeHidden?: boolean },
): Promise<FolderSummary[]> {
  const query: Record<string, string> = {
    $select: 'Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,TotalItemCount',
  };
  if (opts.includeHidden) {
    query['includeHiddenFolders'] = 'true';
  }

  return client.listAll<FolderSummary>(
    `/me/MailFolders/${encodeURIComponent(parentId)}/childfolders`,
    query,
    { top: opts.top ?? 250, maxPages: 50 },
  );
}
```

---

## 6. Recommended Caps

| Cap                                                                     | Recommended Value | Rationale                                                                                                                                                                                                                                                                       |
| ----------------------------------------------------------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `$top` per page                                                         | **250**           | Balances payload size (~50 KB/page) vs. round-trips. Safely below the practical 1000 limit. Server may return fewer; `@odata.nextLink` compensates.                                                                                                                             |
| `maxPages` per collection (`MAX_FOLDER_PAGES`)                          | **50**            | Bounds a single `listChildren` call to at most 50 × 250 = 12,500 folders at one level. No real mailbox has that many direct children; the cap stops runaway loops on pathological data. Maps to `UPSTREAM_PAGINATION_LIMIT` (exit 5) with an actionable message.                |
| Max total nodes visited during a recursive walk (`MAX_FOLDERS_VISITED`) | **5,000**         | Across all levels of a `list-folders --recursive` tree walk. Each node costs one membership in the results array; at ~200 bytes/node this is ~1 MB of in-memory data. Exceeding this raises `UPSTREAM_PAGINATION_LIMIT` exit 5 with guidance to use `--parent` to narrow scope. |
| Max path depth (`MAX_PATH_SEGMENTS`)                                    | **16**            | Already specified in `investigation-folders.md §4.2`. Bounds recursive resolver depth.                                                                                                                                                                                          |

### Sizing reasoning

- A typical Exchange Online mailbox has 15-30 top-level folders and 50-200 total folders
  across all levels.
- The 50-page cap at one level = 12,500 folders before failing; this is a safety net, not
  a normal operating range.
- The 5,000-node walk cap was chosen as: 5,000 × (2 round-trips per level average) =
  10,000 API calls worst case. At 50ms/call that is 8+ minutes — clearly pathological.
  In practice a 3-level tree with 100 siblings/level costs 1 + 1 + 1 calls = 3 calls
  for a single path resolution.

---

## Assumptions & Scope

| Assumption                                                                    | Confidence | Impact if Wrong                                                                                                                                                                 |
| ----------------------------------------------------------------------------- | ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| v2.0 and Graph share the same Exchange Online paging backend                  | HIGH       | No impact in practice; investigation-folders.md already scopes to v2.0 only                                                                                                     |
| Default page size for `mailFolders` is 10                                     | HIGH       | Graph docs explicitly state "default page size (10 items)"; v2.0 inherits the same backend                                                                                      |
| `@odata.nextLink` uses `$skip` (not `$skiptoken`) for Mail folder collections | HIGH       | MS Graph docs explicitly call this out for "Outlook Mail and Calendars". If wrong, the helper's "follow verbatim" strategy still works — it just parses a different query param |
| `DisplayName` is filterable (server-side `$filter`)                           | MEDIUM     | Schema marks it Yes; real-world tenant reliability varies. Client-side fallback is primary strategy regardless                                                                  |
| Max `$top` for `childfolders` is effectively ~1000                            | MEDIUM     | Not officially documented for folders (only for messages). Using 250 keeps a safe margin                                                                                        |
| `@odata.nextLink` is always absolute and on `outlook.office.com`              | HIGH       | OData spec mandates absolute URLs; investigation already mandates host validation                                                                                               |

### Explicitly out of scope

- `$batch` support for folder enumeration (NG4 in refined spec).
- Graph API migration (`/me/mailFolders` camelCase) — future iteration.
- `$skiptoken`-style cursors — Mail/Calendar APIs use `$skip`, not skiptoken.
- `$delta` (folder sync) — separate concern, not needed by the resolver.

---

## References

| #   | Source                                           | URL                                                                                                                                               | Information Gathered                                                                                                                                                                                                                                          |
| --- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1   | Microsoft Graph — List mailFolders               | https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders?view=graph-rest-1.0                                                             | **Confirmed default page size = 10 items** (explicit tip in docs); `@odata.nextLink` emitted when collection > default page size; `includeHiddenFolders` query param                                                                                          |
| 2   | Microsoft Graph — List childFolders              | https://learn.microsoft.com/en-us/graph/api/mailfolder-list-childfolders?view=graph-rest-1.0                                                      | Endpoint shape, `includeHiddenFolders`, OData query params supported, response shape with `isHidden` field                                                                                                                                                    |
| 3   | Microsoft Graph — MailFolder resource            | https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0                                                              | Full well-known folder name list (archive, clutter, conflicts, conversationhistory, deleteditems, drafts, inbox, junkemail, localfailures, msgfolderroot, outbox, recoverableitemsdeletions, scheduled, searchfolders, sentitems, serverfailures, syncissues) |
| 4   | Microsoft Graph — Paging                         | https://learn.microsoft.com/en-us/graph/paging                                                                                                    | **Confirmed `$skip` (not `$skiptoken`) for Outlook Mail/Calendar APIs** explicitly; `@odata.nextLink` must be followed verbatim; DirectoryPageTokenNotFoundException error pattern for retry-token misuse                                                     |
| 5   | Microsoft Graph — Query parameters               | https://learn.microsoft.com/en-us/graph/query-parameters?view=graph-rest-1.0                                                                      | `$top`, `$skip`, `$filter`, `$select` syntax; single-quote escaping rule for `$filter` string values (`O''Brien`)                                                                                                                                             |
| 6   | Microsoft Graph docs-contrib — Paging with $skip | https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/concepts/query-parameters.md                                             | Code snippet context confirming `$skip` paging for Outlook mail/calendar; skipToken used by directory objects only                                                                                                                                            |
| 7   | Outlook REST v2.0 — MailFolder resource schema   | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/complex-types-for-mail-contacts-calendar#FolderResource | **`DisplayName` Filterable: Yes**; `ChildFolderCount`, `TotalItemCount`, `UnreadItemCount` also filterable; `Id` and `ParentFolderId` are NOT filterable                                                                                                      |
| 8   | Outlook REST v2.0 — Mail operations              | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations                                    | Folder operations surface; well-known alias list for v2.0 (`Inbox`, `Drafts`, `SentItems`, `DeletedItems`); sample response showing `@odata.nextLink` with `$skip`                                                                                            |
| 9   | OData v4 spec — Server-Driven Paging             | https://docs.oasis-open.org/odata/odata/v4.01/os/part1-protocol/odata-v4.01-os-part1-protocol.html#sec_ServerDrivenPaging                         | `@odata.nextLink` is an absolute URL the client follows verbatim; server controls the cursor                                                                                                                                                                  |
| 10  | investigation-folders.md (this project)          | docs/design/investigation-folders.md                                                                                                              | Research motivation (§6 Topic 3); sample `@odata.nextLink` showing `$top=100&$skip=100` format confirmed independently                                                                                                                                        |

### Recommended for Deep Reading

- **Source 4 (Paging)**: Clearest single-page explanation of which Graph APIs use `$skip`
  vs. `$skiptoken`, retry-token pitfalls, and the "follow verbatim" rule.
- **Source 7 (v2.0 MailFolder schema)**: The filterable-field table is the authoritative
  answer to whether `$filter=DisplayName` is schema-legal on v2.0.
- **Source 1 (List mailFolders)**: The "Tip" note about default page size = 10 is
  critical for sizing the `$top` default in `listAll<T>`.
