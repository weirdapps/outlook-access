# Refined Request: Outlook CLI — Folder Management (search, create, move, list-in-folder)

## 1. Summary

Extend the already-shipped `outlook-cli` TypeScript tool with folder-management
capabilities on top of Outlook REST v2.0 (`https://outlook.office.com/api/v2.0/me/...`).
The extension introduces three user-facing capabilities and the shared machinery that
backs them:

1. **Search / list / create folders** — enumerate top-level and child mail folders,
   resolve a folder by name or by path (e.g. `Inbox/Projects/Alpha`), and create
   a new folder either directly under the mailbox root, or nested under an
   arbitrary parent.
2. **Move emails to folders** — move one or more messages to a destination folder
   identified by id, well-known folder name, or path. The move endpoint returns a
   _new_ message id in the destination folder; the tool surfaces the id mapping.
3. **List emails in a specific folder** — extend/replace the current `list-mail`
   folder restriction (today: a closed set of well-known names) so that any user
   folder reachable by name or path can be used, while preserving backward
   compatibility with the existing `--folder <well-known>` flag.

The extension reuses the existing auth, session, HTTP, output, and error layers
and must NOT relax the project-wide rules: no fallback defaults for mandatory
configuration, secrets never logged, session file atomic + `0600`, exit codes
aligned with the current 0/1/2/3/4/5/6 taxonomy.

## 2. Goals

- **G1.** Add three new subcommands (`list-folders`, `find-folder`, `create-folder`)
  plus one new subcommand for moving mail (`move-mail`), and extend the existing
  `list-mail` to accept an arbitrary folder by id or path while remaining
  backward-compatible with the current well-known name set.
- **G2.** Provide a single canonical **folder resolver** (new module
  `src/folders/resolver.ts`) used by every folder-aware command so that name/path
  semantics, escaping rules, and well-known alias handling are identical across
  the CLI.
- **G3.** Preserve the existing JSON/table output duality. Every new command
  supports `--json` (default) and `--table`.
- **G4.** Produce deterministic, scriptable JSON payloads with a schema that
  downstream scripts can depend on (documented in §5 per subcommand).
- **G5.** Reuse the existing `OutlookClient`, auth flow, 401-retry-once logic,
  and error classes (`ConfigurationError` / `AuthError` / `UpstreamError` /
  `UsageError`) — no new exit-code values introduced.
- **G6.** Update `CLAUDE.md` `<outlook-cli>` documentation with the new
  subcommands, flags, exit codes, and examples in the same `<toolName>` block
  already in use.
- **G7.** Update `docs/design/project-design.md` and `docs/design/project-functions.MD`
  to reflect the new feature set and module layout.

## 3. Non-Goals (Deferred)

- **NG1.** **Renaming** folders (`PATCH /MailFolders/{id}` DisplayName). Out of
  scope; may be added as a later `rename-folder` subcommand.
- **NG2.** **Deleting** folders (`DELETE /MailFolders/{id}`). Out of scope —
  destructive, needs additional confirmation UX.
- **NG3.** **Copying** messages (`POST /messages/{id}/copy`). Move only.
- **NG4.** **Batch move** across multiple messages in a single REST request
  (`$batch`). Multi-message moves are serial, one REST call per message.
- **NG5.** Subscribing to folder changes (push/delta/websocket).
- **NG6.** **Search folders** (the Outlook server-side virtual folder type) —
  creation and management. Still surfaced read-only if encountered during list.
- **NG7.** Changing a folder's **parent** (move-folder). Only message move.
- **NG8.** Setting **IsHidden**, **ChildFolderCount** recomputation, or any
  other folder property beyond `DisplayName` on create.
- **NG9.** Search-folder / shared-mailbox / archive-mailbox access. Only the
  signed-in user's primary mailbox folders (`/me/MailFolders`).
- **NG10.** Concurrency/parallel move. One REST call at a time.

## 4. User Stories

- **US1. list-folders** — As a user, I run `outlook-cli list-folders` to get
  the top-level mail folders in my mailbox, or `outlook-cli list-folders --parent Inbox --recursive`
  to enumerate the full sub-tree under Inbox.
- **US2. find-folder** — As a user, I run `outlook-cli find-folder "Projects/Alpha"`
  to resolve a path into a folder id (and full metadata) so I can pipe it into
  other commands. If the path is ambiguous (two sibling folders with the same
  name), I get a clear error telling me so.
- **US3. create-folder** — As a user, I run
  `outlook-cli create-folder "Projects/Alpha"` to create a new folder under
  Inbox/Projects, creating intermediate parents with `--create-parents` if needed.
  If the folder already exists and `--idempotent` is set, I get success with the
  existing folder's id.
- **US4. move-mail** — As a user, I run
  `outlook-cli move-mail AAMkAGI... --to "Projects/Alpha"` to move a single
  message to the destination folder. Or I pipe a list of ids via `--ids-from -`
  (stdin) or `--ids-from ids.txt` to move many messages serially.
- **US5. list-mail extended** — As a user, I run
  `outlook-cli list-mail --folder "Projects/Alpha"` or
  `outlook-cli list-mail --folder-id AAMkAGI...` to list messages in any
  user-created folder, not just the five well-known ones currently supported.
- **US6. Pipelining** — As a scripter, I chain commands:
  `outlook-cli find-folder "Projects/Alpha" --json | jq -r .id` then use that id
  with `move-mail --to-id` or `list-mail --folder-id`. Every command that
  consumes a folder accepts both the human form and the id form.

## 5. CLI Surface (new and changed commands)

### 5.0 Global flags (unchanged — reused from existing tool)

All existing global flags (`--timeout`, `--login-timeout`, `--chrome-channel`,
`--session-file`, `--profile-dir`, `--tz`, `--json`, `--table`, `--quiet`,
`--no-auto-reauth`, `--log-file`) apply to every new subcommand unchanged.
Mandatory-env rules apply (see existing §8 of `refined-request-outlook-cli.md`):
`OUTLOOK_CLI_HTTP_TIMEOUT_MS`, `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`, and
`OUTLOOK_CLI_CHROME_CHANNEL` remain mandatory and still cause exit 3 if absent
and not supplied via their flags.

### 5.1 `outlook-cli list-folders`

- **Arguments:** none.
- **Options:**
  - `--parent <name-or-path-or-id>` — optional. Limits enumeration to children
    of the given parent. Accepts:
    - a **well-known alias** (see §6.2): `Inbox`, `SentItems`, `Drafts`,
      `DeletedItems`, `Archive`, `JunkEmail`, `Outbox`, `MsgFolderRoot`,
    - a **display-name path** (see §6.1) such as `Inbox/Projects`,
    - a **folder id** (opaque base64-ish token) when prefixed with `id:` —
      e.g. `--parent id:AAMkAGI...`.
    - Default (when omitted): `MsgFolderRoot` (top-level folders in the
      mailbox).
  - `--recursive` — if set, emit the full sub-tree under the parent. Default:
    only direct children.
  - `--include-hidden` — if set, include folders with `IsHidden == true`.
    Default: `false`.
  - `--top <N>` — upper bound for `$top` passed to each `/childfolders` page.
    Default: `100`. Range: `1..250`.
- **REST targets:**
  - `GET /api/v2.0/me/MailFolders?$top=N&$select=...` for the root level.
  - `GET /api/v2.0/me/MailFolders/{id}/childfolders?$top=N&$select=...` for
    children.
  - Pagination via `@odata.nextLink` handled inside the HTTP layer (follow up
    to an upper bound defined in §7; documented here).
- **Default `$select` fields:** `Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,TotalItemCount,WellKnownName`
  (note: `WellKnownName` is only populated by Outlook on well-known folders.)
- **Output (JSON):** an array of `FolderSummary` objects:

  ```json
  [
    {
      "Id": "AAMkAGI...",
      "DisplayName": "Inbox",
      "ParentFolderId": "AAMkAGI...root...",
      "ChildFolderCount": 3,
      "UnreadItemCount": 12,
      "TotalItemCount": 402,
      "WellKnownName": "inbox",
      "Path": "Inbox"
    }
  ]
  ```

  When `--recursive` is set, `Path` is the materialized display-name path from
  the parent down (using `/` separator — see §6.1).

- **Output (table):** columns `Path | Unread | Total | Children | Id`.
- **Exit codes:** 0; 2 (bad flag); 3 (missing mandatory config); 4; 5; 6 (n/a
  here, listed for symmetry).

### 5.2 `outlook-cli find-folder <query>`

- **Arguments:**
  - `<query>` — **required** positional. Accepted forms:
    - well-known alias (`Inbox`, `Archive`, …),
    - display-name path (`Inbox/Projects/Alpha`),
    - id form (`id:AAMkAGI...`) — in which case `find-folder` degenerates into
      a single GET + echo (useful for normalization).
- **Options:**
  - `--parent <name-or-path-or-id>` — when `<query>` is a **bare name** (not a
    path), resolve within this parent. Defaults to `MsgFolderRoot`.
  - `--case-sensitive` — if set, compare DisplayName case-sensitively (default:
    case-insensitive, Unicode-aware — see §6.3).
  - `--include-hidden` — include `IsHidden == true` during lookup.
- **REST target:**
  - For a path, walk children level by level using
    `GET /api/v2.0/me/MailFolders/{parentId}/childfolders?$filter=DisplayName eq '<escaped>'`.
  - For id form, `GET /api/v2.0/me/MailFolders/{id}`.
- **Output (JSON):** a single `FolderSummary` (same shape as §5.1 element)
  plus a `ResolvedVia` field: `"wellknown" | "path" | "id"`.
- **Output (table):** two-line key/value display for each field.
- **Errors / exit codes:**
  - Path component not found → `UpstreamError` code `UPSTREAM_FOLDER_NOT_FOUND`,
    exit 5. (Rationale: this is a "not present on the server" outcome, not a
    usage error; symmetric with `get-mail` on a bogus id.)
  - Ambiguous path (two siblings share the same DisplayName) → `UsageError`
    code `FOLDER_AMBIGUOUS`, exit 2. Error payload lists the matching ids.
  - Id form for a non-existent id → `UpstreamError` code
    `UPSTREAM_FOLDER_NOT_FOUND`, exit 5.

### 5.3 `outlook-cli create-folder <path>`

- **Arguments:**
  - `<path>` — **required** positional. Must be a display-name path (`A/B/C`)
    or a bare name (treated as path length 1). Well-known aliases are NOT
    valid create targets (cannot create a folder _named_ `Inbox` at the root —
    would collide with the well-known one; see §6.2).
- **Options:**
  - `--parent <name-or-path-or-id>` — anchor for the path. Default:
    `MsgFolderRoot` (i.e. top-level). If set, the new folder(s) are created
    relative to this parent.
  - `--create-parents` — if set, intermediate missing parents along `<path>`
    are created. If not set, any missing intermediate parent → `UsageError`
    exit 2.
  - `--idempotent` — if set, when the final leaf folder already exists under
    the computed parent, exit 0 and return its existing id. Without
    `--idempotent`, exit 6 (IO-like collision) with code `FOLDER_ALREADY_EXISTS`.
  - `--display-name <name>` — optional override of the last path segment's
    display name (useful if the path contains characters that must be path-
    escaped but the actual DisplayName on the server should not be). If
    omitted, the last path segment is used verbatim (after unescape — §6.1).
- **REST target:** `POST /api/v2.0/me/MailFolders/{parentId}/childfolders`
  with body `{"DisplayName": "<name>"}`. For nested creation, the call is
  repeated once per missing parent, then once for the leaf.
- **Output (JSON):**

  ```json
  {
    "created": [
      {
        "Id": "AAMkAGI...Projects",
        "DisplayName": "Projects",
        "Path": "Projects",
        "ParentFolderId": "AAMkAGI...root",
        "PreExisting": false
      },
      {
        "Id": "AAMkAGI...Alpha",
        "DisplayName": "Alpha",
        "Path": "Projects/Alpha",
        "ParentFolderId": "AAMkAGI...Projects",
        "PreExisting": false
      }
    ],
    "leaf": { "Id": "AAMkAGI...Alpha", "Path": "Projects/Alpha", "DisplayName": "Alpha" },
    "idempotent": false
  }
  ```

  When `--idempotent` matches an existing leaf: `created` is empty,
  `leaf.PreExisting = true`, top-level `idempotent` is `true`.

- **Output (table):** columns `Path | Id | PreExisting`.
- **Exit codes:** 0; 2 (missing `<path>`, `/`-only path, invalid chars, or
  missing intermediate parent without `--create-parents`); 3; 4; 5
  (`POST /childfolders` fails with 4xx/5xx); 6 (leaf collision without
  `--idempotent`). `FOLDER_ALREADY_EXISTS` maps to exit 6 because the cause
  is "state on disk/server already occupies this slot", which matches the
  existing exit-6 semantics used by `download-attachments` for filename
  collisions.

### 5.4 `outlook-cli move-mail <id>`

- **Arguments:**
  - `<id>` — **optional**, positional, single message id. Mutually exclusive
    with `--ids-from`. If both absent → `UsageError` exit 2. If both present →
    `UsageError` exit 2.
- **Options:**
  - `--to <name-or-path>` — destination folder expressed as a well-known
    alias or as a display-name path. Exactly one of `--to` / `--to-id` is
    required; otherwise → `UsageError` exit 2.
  - `--to-id <folderId>` — destination folder expressed as an id (skips the
    resolver, goes straight to the move call).
  - `--to-parent <name-or-path-or-id>` — when `--to` is a bare name, anchor
    its resolution to this parent. Default: `MsgFolderRoot`.
  - `--ids-from <path-or-dash>` — read ids from a file (one id per line,
    blank lines and `#` comments ignored), or from stdin if `-` is given.
  - `--continue-on-error` — if set, a failure on one id does not abort the
    whole run. Default: abort on first failure.
  - `--stop-at <N>` — upper bound on the number of ids processed in a single
    run when using `--ids-from`. Default: `1000`. Range: `1..10000`. Guards
    against accidental mailbox-wide moves.
- **REST target:** `POST /api/v2.0/me/messages/{id}/move` with body
  `{"DestinationId": "<resolvedId>"}`. The server returns the **new** message
  resource (with a **new Id** because ImmutableId is not used in REST v2.0
  defaults) in the destination folder.
- **Output (JSON):**

  ```json
  {
    "destination": { "Id": "AAMkAGI...dest", "Path": "Projects/Alpha", "DisplayName": "Alpha" },
    "moved": [
      { "sourceId": "AAMkAGI...srcA", "newId": "AAMkAGI...newA" },
      { "sourceId": "AAMkAGI...srcB", "newId": "AAMkAGI...newB" }
    ],
    "failed": [
      { "sourceId": "AAMkAGI...srcC", "error": { "code": "UPSTREAM_HTTP_404", "httpStatus": 404 } }
    ],
    "summary": { "requested": 3, "moved": 2, "failed": 1 }
  }
  ```

  (The `failed` array is populated only when `--continue-on-error` is in effect
  or when `--ids-from` semantics apply — see below.)

- **Output (table):** columns `Source Id | New Id | Status | Error`.
- **Exit codes / semantics:**
  - 0 — all requested ids moved successfully.
  - 2 — usage error (missing id/`--ids-from`, both present, both `--to`/`--to-id`,
    missing both).
  - 3 — missing mandatory config.
  - 4 — auth failure.
  - 5 — upstream error on the destination resolve or on any move that was NOT
    absorbed by `--continue-on-error`.
  - 6 — reserved; unused by `move-mail` in this iteration.
  - When `--continue-on-error` is set and at least one id fails but at least
    one succeeded → exit **5** (an upstream failure occurred) — but the
    `failed[]` array is still fully populated in the JSON output so scripts
    can re-drive the failures. Rationale: silently returning 0 would hide
    partial failure; using a new exit code breaks the existing taxonomy.

### 5.5 `outlook-cli list-mail` — extension

**Backward-compatibility:** the existing `--folder <WellKnownName>` flag
continues to work exactly as before — the current hardcoded set `Inbox`,
`SentItems`, `Drafts`, `DeletedItems`, `Archive` remains valid input.

**New behaviour (additive):**

- `--folder` now also accepts:
  - any other **well-known alias** recognized by Outlook: `JunkEmail`,
    `Outbox`, `MsgFolderRoot`, `RecoverableItemsDeletions` (read-only). See
    §6.2.
  - a **display-name path** (`Inbox/Projects/Alpha`) — resolved via the folder
    resolver.
- **New flag:** `--folder-id <id>` — mutually exclusive with `--folder`. Skips
  the resolver, uses the id directly. If both flags are present → `UsageError`
  exit 2. If neither is present, the existing default (`Inbox`) applies.
- **New flag:** `--folder-parent <name-or-path-or-id>` — anchor for a bare
  `--folder` name that is not a well-known alias. Default: `MsgFolderRoot`.
- **URL construction:**
  - When the effective folder is resolved to an id, the request becomes
    `GET /api/v2.0/me/MailFolders/{id}/messages?...`.
  - When the effective folder is a well-known alias, the request remains the
    existing form `GET /api/v2.0/me/MailFolders/<alias>/messages?...`
    (Outlook accepts the alias in the path).
- **Output (JSON):** unchanged — still `MessageSummary[]`.
- **Output (table):** unchanged — still `Received | From | Subject | Att | Id`.
- **Exit codes:** unchanged; ambiguous resolution and not-found map to the
  same codes as `find-folder` (§5.2).

## 6. Folder naming, path, and alias rules (normative)

### 6.1 Path syntax

- **Separator:** `/` (forward slash) between display-name segments.
- **Escaping:**
  - A literal `/` inside a DisplayName is encoded as `\/`.
  - A literal `\` inside a DisplayName is encoded as `\\`.
  - No other characters require escaping in paths (leading/trailing whitespace
    is preserved, not trimmed).
  - Rationale: Outlook DisplayName allows `/`, so we cannot use `/` as a pure
    separator without escaping. `\` is the natural escape character.
- **Normalization:** segments are NOT lowercased, NFC-normalized for
  comparison on the client side (Unicode NFC; see §6.3).
- **Empty segments forbidden:** `//` or leading/trailing `/` (after the
  parent anchor) → `UsageError` exit 2.
- **Maximum path depth:** 16 segments. Paths deeper than that → `UsageError`
  exit 2. (Outlook supports deeper in principle; 16 is a safety cap for
  client-side walk logic.)

### 6.2 Well-known folder aliases

The tool recognizes the following aliases — these resolve at the **Outlook
server level** without a lookup call (Outlook accepts the literal token in the
URL path):

| Alias                       | Notes                                                 |
| --------------------------- | ----------------------------------------------------- |
| `Inbox`                     | Read/write                                            |
| `SentItems`                 | Read/write                                            |
| `Drafts`                    | Read/write                                            |
| `DeletedItems`              | Read/write                                            |
| `Archive`                   | Read/write (not guaranteed on every tenant)           |
| `JunkEmail`                 | Read/write                                            |
| `Outbox`                    | Read only in practice                                 |
| `MsgFolderRoot`             | The mailbox root; **container for top-level folders** |
| `RecoverableItemsDeletions` | Read only                                             |

**Precedence when a user folder has the same DisplayName as a well-known
alias:** the well-known alias wins at the **root** (e.g. `--folder Inbox`
without a `--folder-parent` always resolves to the well-known Inbox, never to
a user-created top-level folder named `Inbox`). Inside any non-root parent,
well-known aliases are NOT matched — a path like `Inbox/Inbox` resolves by
display-name lookup inside the real Inbox. This rule is documented in
`find-folder --help` output.

### 6.3 Name matching

- **Equality comparison** uses Unicode NFC + case-folding (Unicode simple
  case fold) by default. `--case-sensitive` disables case-folding (still
  applies NFC).
- **No trimming, no collapsing of internal whitespace.** A folder whose
  DisplayName has trailing whitespace must be addressed with that whitespace
  in the path segment.
- **`$filter` encoding:** single quotes in DisplayName are doubled for OData
  (`DisplayName eq 'O''Brien'`) by the HTTP layer. The tool builds the filter
  string server-side-safely; callers must NOT pre-escape.

### 6.4 Ambiguity policy

- Two siblings can share the same DisplayName (Outlook allows it). `find-folder`
  and any resolver invocation that encounters multiple matches at a given
  level raises `UsageError` `FOLDER_AMBIGUOUS` (exit 2) unless an id form is
  used.
- **Disambiguation knob:** `--first-match` flag on `find-folder`, `list-mail`,
  `move-mail` selects the **first** match by `CreatedDateTime asc` (stable
  tiebreaker: `Id` lexicographic asc). Documented as a foot-gun; not the
  default.

## 7. HTTP, pagination, rate-limit, retry

- **Pagination.** Every list-type folder call honors `@odata.nextLink` up to
  an upper bound of 50 pages. Beyond 50 pages an `UpstreamError` code
  `UPSTREAM_PAGINATION_LIMIT` (exit 5) is raised.
- **401 retry-once.** Identical to existing commands: on 401, trigger a
  single re-auth via `onReauthNeeded`, rebuild the client, retry once. Second
  401 → `AuthError` `AUTH_401_AFTER_RETRY`, exit 4.
- **Timeouts.** Each HTTP call uses the mandatory `--timeout` value; no
  per-command override. Resolver walks issue one request per depth level —
  each is bounded by the same per-call timeout.
- **Move side-effects.** `POST /messages/{id}/move` returns **200** with the
  new message (and new Id). On **404** for the source id, the tool does NOT
  re-auth; it maps to `UPSTREAM_HTTP_404`, exit 5 (or absorbs into `failed[]`
  when `--continue-on-error` is set).
- **Create idempotency on the wire.** The server returns **201 Created** with
  the new folder body. If a folder with the same DisplayName exists under
  the same parent, Outlook returns **409** (or **400** on some tenants). The
  resolver must treat both as "already exists" when `--idempotent` is set,
  then issue a lookup to return the existing id. Without `--idempotent` the
  original upstream status is preserved and the tool exits 6 with
  `FOLDER_ALREADY_EXISTS`.

## 8. Module layout (additive to existing `src/`)

New:

- `src/folders/resolver.ts`
  - `resolveFolder(client, spec: FolderSpec, opts): Promise<ResolvedFolder>`
  - `listChildren(client, parentId, opts): Promise<FolderSummary[]>`
  - `createFolderPath(client, path, opts): Promise<CreateResult>`
  - `parseFolderPath(input: string): string[]` (segmenter + unescape)
  - `buildFolderPath(segments: string[]): string`
  - `matchesWellKnownAlias(input: string): string | null`
- `src/folders/types.ts`
  - `FolderSummary`, `FolderSpec` (`{ kind: 'wellknown' | 'path' | 'id', value: string, parent?: FolderSpec }`),
    `ResolvedFolder`, `CreateResult`, `MoveResult`.
- `src/commands/list-folders.ts`
- `src/commands/find-folder.ts`
- `src/commands/create-folder.ts`
- `src/commands/move-mail.ts`

Changed:

- `src/commands/list-mail.ts` — accept `--folder-id`, accept path/non-well-
  known names, delegate to resolver when needed. Must remain backward-
  compatible for the existing well-known subset.
- `src/cli.ts` — wire the four new subcommands and the two new flags on
  `list-mail`.
- `src/http/outlook-client.ts` — add `client.post(path, body)` if not already
  present (needed for `/move` and `/childfolders`), and a generic list helper
  that follows `@odata.nextLink`.
- `src/http/errors.ts` — extend `ApiError.code` vocabulary with
  `UPSTREAM_FOLDER_NOT_FOUND`, `UPSTREAM_FOLDER_AMBIGUOUS` (raised at the
  resolver layer but re-classified), `UPSTREAM_PAGINATION_LIMIT`. No new
  class.
- `src/config/errors.ts` — extend `UsageError.code` vocabulary with
  `FOLDER_AMBIGUOUS`, `FOLDER_MISSING_PARENT`, `FOLDER_PATH_INVALID`. Add
  `UpstreamError.code` `FOLDER_ALREADY_EXISTS` — mapped to exit 6 via a small
  override in the top-level handler (or, cleaner, a new dedicated
  `CollisionError` with exit 6, mirroring the existing attachment-collision
  exit). Implementation phase chooses one.
- `src/output/formatter.ts` — add table formatters for `FolderSummary[]`,
  `CreateResult`, `MoveResult`.

## 9. Configuration additions

No new **mandatory** configuration variables. All behavior is flag-driven or
derived from existing mandatory flags (timeouts, Chrome channel).

New **optional** flags on existing commands:

| Command       | Flag                  | Default         | Env var (optional mirror) |
| ------------- | --------------------- | --------------- | ------------------------- |
| list-mail     | `--folder-id`         | (unset)         | —                         |
| list-mail     | `--folder-parent`     | `MsgFolderRoot` | —                         |
| list-folders  | `--parent`            | `MsgFolderRoot` | —                         |
| list-folders  | `--recursive`         | `false`         | —                         |
| list-folders  | `--include-hidden`    | `false`         | —                         |
| list-folders  | `--top`               | `100`           | —                         |
| find-folder   | `--parent`            | `MsgFolderRoot` | —                         |
| find-folder   | `--case-sensitive`    | `false`         | —                         |
| find-folder   | `--include-hidden`    | `false`         | —                         |
| find-folder   | `--first-match`       | `false`         | —                         |
| create-folder | `--parent`            | `MsgFolderRoot` | —                         |
| create-folder | `--create-parents`    | `false`         | —                         |
| create-folder | `--idempotent`        | `false`         | —                         |
| create-folder | `--display-name`      | (last segment)  | —                         |
| move-mail     | `--to` / `--to-id`    | (required)      | —                         |
| move-mail     | `--to-parent`         | `MsgFolderRoot` | —                         |
| move-mail     | `--ids-from`          | (unset)         | —                         |
| move-mail     | `--continue-on-error` | `false`         | —                         |
| move-mail     | `--stop-at`           | `1000`          | —                         |
| move-mail     | `--first-match`       | `false`         | —                         |

The project-wide rule — **no fallback defaults for mandatory settings** — is
respected: every default above is for a non-mandatory flag, and each default
is explicitly enumerated in this spec (matching the existing `refined-request-
outlook-cli.md` §8 policy).

## 10. Error classes and exit-code mapping

| Scenario                                                        | Error class / code                                | Exit |
| --------------------------------------------------------------- | ------------------------------------------------- | ---- |
| Missing `<query>` on `find-folder`                              | `UsageError` / `BAD_USAGE`                        | 2    |
| Missing `<path>` on `create-folder`                             | `UsageError` / `BAD_USAGE`                        | 2    |
| Both `<id>` and `--ids-from` on `move-mail`                     | `UsageError` / `BAD_USAGE`                        | 2    |
| Neither `--to` nor `--to-id` on `move-mail`                     | `UsageError` / `BAD_USAGE`                        | 2    |
| Both `--folder` and `--folder-id` on `list-mail`                | `UsageError` / `BAD_USAGE`                        | 2    |
| Path with empty segment / invalid escape                        | `UsageError` / `FOLDER_PATH_INVALID`              | 2    |
| Missing intermediate parent, no `--create-parents`              | `UsageError` / `FOLDER_MISSING_PARENT`            | 2    |
| Ambiguous path match, no `--first-match`                        | `UsageError` / `FOLDER_AMBIGUOUS`                 | 2    |
| Mandatory config absent (`--timeout`, `--chrome-channel`, etc.) | `ConfigurationError`                              | 3    |
| Interactive login cancelled or timed out                        | `AuthError`                                       | 4    |
| 401 from folder / move / create after retry                     | `AuthError` / `AUTH_401_AFTER_RETRY`              | 4    |
| 404 on folder lookup (path, id)                                 | `UpstreamError` / `UPSTREAM_FOLDER_NOT_FOUND`     | 5    |
| 404 on source message during move                               | `UpstreamError` / `UPSTREAM_HTTP_404`             | 5    |
| Any other 4xx/5xx from folder/move/create                       | `UpstreamError` / `UPSTREAM_HTTP_<status>`        | 5    |
| Network / timeout / abort                                       | `UpstreamError` / `UPSTREAM_NETWORK` / `_TIMEOUT` | 5    |
| Pagination cap exceeded                                         | `UpstreamError` / `UPSTREAM_PAGINATION_LIMIT`     | 5    |
| Folder already exists, no `--idempotent`                        | `FOLDER_ALREADY_EXISTS` (see §8 for class choice) | 6    |

## 11. Acceptance Criteria

Each AC below must be backed by at least one test script under `test_scripts/`
(unit + integration-with-fake-client where the real API is not available).

### Passing scenarios

1. **AC-LISTFOLDERS-ROOT** — `outlook-cli list-folders` returns a non-empty
   array including `Inbox` with `WellKnownName == "inbox"`.
2. **AC-LISTFOLDERS-CHILDREN** — `outlook-cli list-folders --parent Inbox`
   returns the direct children of Inbox. With `--recursive`, the sub-tree is
   flattened and `Path` is populated (`Inbox/Projects`, `Inbox/Projects/Alpha`, …).
3. **AC-FIND-WELLKNOWN** — `outlook-cli find-folder Inbox` returns the
   well-known Inbox with `ResolvedVia: "wellknown"`.
4. **AC-FIND-PATH** — `outlook-cli find-folder "Projects/Alpha" --parent Inbox`
   resolves the nested folder by path and returns its id.
5. **AC-FIND-ID** — `outlook-cli find-folder id:AAMkAGI...` normalizes and
   returns the folder metadata with `ResolvedVia: "id"`.
6. **AC-CREATE-TOPLEVEL** — `outlook-cli create-folder "Test-$(date +%s)"`
   creates a new top-level folder; output shows `created[0].PreExisting == false`
   and `leaf.Id` is non-empty.
7. **AC-CREATE-NESTED** — `outlook-cli create-folder "A/B/C" --parent Inbox
--create-parents` creates (or reuses) `Inbox/A`, `Inbox/A/B`, `Inbox/A/B/C`.
8. **AC-CREATE-IDEMPOTENT** — Re-running AC-CREATE-NESTED with `--idempotent`
   exits 0; `created[]` is empty, `idempotent == true`.
9. **AC-MOVE-SINGLE** — `outlook-cli move-mail <id> --to "Inbox/Projects/Alpha"`
   returns `moved[0].newId != moved[0].sourceId`; the source id is no longer
   retrievable in the source folder.
10. **AC-MOVE-MANY** — `outlook-cli move-mail --ids-from ids.txt --to-id <id>`
    moves every id; summary matches the count in the file.
11. **AC-LISTMAIL-PATH** — `outlook-cli list-mail --folder "Inbox/Projects/Alpha"
-n 5` returns 5 messages from that folder.
12. **AC-LISTMAIL-ID** — `outlook-cli list-mail --folder-id AAMkAGI... -n 5`
    returns 5 messages (same IDs regardless of how the folder is addressed).
13. **AC-LISTMAIL-WELLKNOWN-BACKCOMPAT** — `outlook-cli list-mail --folder Inbox
-n 3` continues to work exactly as before (no resolver hop).

### Failing / edge scenarios

1. **AC-FOLDER-NOT-FOUND** — `outlook-cli find-folder "DoesNotExist"` exits 5
   with `code == "UPSTREAM_FOLDER_NOT_FOUND"`.
2. **AC-FOLDER-AMBIGUOUS** — With two sibling folders named `Projects`,
   `find-folder "Projects"` exits 2 with `code == "FOLDER_AMBIGUOUS"` and
   lists the candidate ids. With `--first-match`, exit 0 and only the first
   is returned.
3. **AC-CREATE-COLLISION** — Creating an existing folder without
   `--idempotent` exits 6 with `code == "FOLDER_ALREADY_EXISTS"`.
4. **AC-CREATE-MISSING-PARENT** — `create-folder "A/B"` without
   `--create-parents` when `A` does not exist exits 2 with
   `code == "FOLDER_MISSING_PARENT"`.
5. **AC-MOVE-BAD-DEST** — `move-mail <id> --to "Nope/Nope"` exits 5 with
   `code == "UPSTREAM_FOLDER_NOT_FOUND"` (or exit 2 `FOLDER_AMBIGUOUS` if
   path is malformed — consistent with §10 table).
6. **AC-MOVE-BAD-SOURCE** — `move-mail bogus --to Inbox` exits 5 with
   `code == "UPSTREAM_HTTP_404"`.
7. **AC-MOVE-PARTIAL** — With `--continue-on-error`, a batch of 3 ids where
   one source id is bogus exits **5** (not 0), but `moved[]` has 2 entries
   and `failed[]` has 1.
8. **AC-MOVE-STOPAT** — With `--ids-from` containing 5 ids and `--stop-at 2`,
   exactly 2 are attempted; the run exits 2 with `code == "BAD_USAGE"` and a
   summary of what would have been moved. (Rationale: `--stop-at` is a safety
   valve; exceeding it is a usage error, not a partial success.)
9. **AC-PATH-ESCAPE** — A folder named `A/B` can be created by passing
   `"A\/B"` on the CLI, and `find-folder "A\/B"` resolves it without
   descending into a non-existent parent `A`.
10. **AC-PATH-DEPTH-CAP** — A path with 17 segments exits 2 with
    `code == "FOLDER_PATH_INVALID"`.
11. **AC-WELLKNOWN-PRECEDENCE** — Given a user-created top-level folder named
    `Inbox`, `find-folder Inbox` resolves to the well-known Inbox. Reaching
    the user one requires `--parent MsgFolderRoot --first-match` or a
    disambiguating path.
12. **AC-401-RETRY-FOLDERS** — A 401 on `POST /childfolders` triggers exactly
    one re-auth, then succeeds. A second 401 exits 4.
13. **AC-NO-SECRET-LEAK-FOLDERS** — With `--log-file` at debug, no folder or
    move log line contains the bearer token or cookie values.
14. **AC-CLAUDEMD-UPDATED-FOLDERS** — `CLAUDE.md`'s `<outlook-cli>` entry
    lists every new subcommand with the exact flag syntax, defaults, and
    exit-code table described here.

## 12. Assumptions

- **A1.** Outlook REST v2.0 remains the API surface; no migration to Graph is
  part of this work.
- **A2.** The `POST /messages/{id}/move` endpoint is available on the tenants
  this tool targets; Outlook web uses it today, so the captured bearer token
  has the necessary scopes.
- **A3.** `POST /MailFolders/{id}/childfolders` creates a folder and returns
  the created resource in the response body (as documented for REST v2.0).
- **A4.** `DisplayName eq '...'` `$filter` on `/childfolders` is supported for
  name lookups. If a specific tenant rejects the filter, the resolver falls
  back to client-side filtering after paginating the children (detected at
  runtime, one extra round-trip). This fallback is an **implementation**
  detail and is intentionally not an acceptance criterion.
- **A5.** Moving a draft or an item in a system-protected folder can return
  4xx; these surface as `UpstreamError` with the upstream status — no
  special-casing.
- **A6.** The project-wide session, auth, and output layers are stable and
  are not refactored as part of this change — only additive extension points
  are touched.
- **A7.** Concurrency is irrelevant for this iteration: all commands run in a
  single process, serially, with no shared mutable state beyond the session
  file (already lock-protected).

## 13. Open Questions

1. **Collision class choice.** Should "folder already exists" use a new
   dedicated `CollisionError` (exit 6) class, or should we extend
   `UpstreamError` with a code and route it to exit 6 via a special case in
   `cli.ts`? (See §8 note.) Default assumption: a new `CollisionError` for
   symmetry with the attachment-collision flow.
2. **`--first-match` tiebreaker.** Is `CreatedDateTime asc` the right
   deterministic tiebreaker, or should it be `DisplayName asc` + `Id asc`?
   Default assumption: `CreatedDateTime asc, Id asc`.
3. **`MsgFolderRoot` parent default.** Should the default parent for
   `create-folder` and `list-folders` be `MsgFolderRoot` (mailbox root,
   sibling to Inbox), or `Inbox` (inside Inbox)? The spec above picks
   `MsgFolderRoot` because that is how Outlook web behaves when the user
   clicks "New folder" at the mailbox root, but confirm this.
4. **Immutable IDs.** Should the move operation opt into the Outlook
   "ImmutableIds" header so that `sourceId == newId` after move? Default
   assumption: **no**, to stay consistent with the existing `list-mail` /
   `get-mail` id semantics. If yes, a global `--immutable-ids` switch is the
   cleanest knob.
5. **`--ids-from` file format.** Plain one-id-per-line is assumed. Should
   JSON-array input also be supported for symmetry with the JSON output of
   `list-mail`? Default assumption: **no**, but easy to add via
   `--ids-format json|lines`.

## 14. Original Request

```text
I want you to add support to search and create folders, move emails to folders, list emails in folders
```
