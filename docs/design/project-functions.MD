# Project Functions — Outlook CLI

This document registers every functional requirement and feature of the Outlook CLI.
Each command exposed by the `outlook-cli` binary is represented as a numbered functional
requirement (FR-xxx). Source of truth: `docs/design/refined-request-outlook-cli.md`.

Last updated: 2026-04-21

---

## FR-001 — `outlook-cli login`

**Purpose.** Open a headed Chrome window via Playwright, let the user complete Microsoft
365 login against `outlook.office.com`, capture the first `Authorization: Bearer <JWT>`
header and the browser cookie jar, and persist the session to
`$HOME/.outlook-cli/session.json` (mode `0600`).

**Inputs.** `--force` (ignore any cached session and always open the browser).
**Outputs (JSON).** `{ status, sessionFile, tokenExpiresAt, account: { puid, tenantId, upn } }`.
**Exit codes.** `0` success, `3` missing mandatory config, `4` login not completed / user
cancelled, `6` IO error writing the session file.
**Source.** Refined request §5.1.

---

## FR-002 — `outlook-cli auth-check`

**Purpose.** Verify that the cached session file exists, its Bearer token has not
expired, and a cheap REST call against `/api/v2.0/me` succeeds. Never launches a browser;
never triggers re-auth.

**Inputs.** none.
**Outputs (JSON).** `{ status: "ok" | "expired" | "missing" | "rejected",
tokenExpiresAt, account: { upn } }`.
**Exit codes.** `0` success (including the `expired`/`missing`/`rejected` reports — they
are informational, not errors), `3` missing mandatory config, `5` upstream error.
**Source.** Refined request §5.2.

---

## FR-003 — `outlook-cli list-mail`

**Purpose.** List the user's N most recent messages in a named folder (default `Inbox`),
ordered by `ReceivedDateTime desc`.

**Inputs.** `-n, --top <N>` (default 10, range 1..100), `--folder <name>` (default
`Inbox`; fast-path well-known set: `Inbox`, `SentItems`, `Drafts`, `DeletedItems`,
`Archive`; also accepts any other well-known alias — `JunkEmail`, `Outbox`,
`MsgFolderRoot`, `RecoverableItemsDeletions` — and arbitrary display-name paths such as
`Inbox/Projects/Alpha`), `--folder-id <id>` (mutually exclusive with `--folder`; skips
the resolver and calls `/me/MailFolders/{id}/messages` directly), `--folder-parent
<name-or-path-or-id>` (default `MsgFolderRoot`; anchor for non-well-known `--folder`
values), `--select <comma-list>` (override selected fields).
**REST target.** `GET /api/v2.0/me/MailFolders/<folder-or-id>/messages?$top=N&$orderby=ReceivedDateTime desc&$select=...`.
**Outputs (JSON).** Array of message summaries (default fields: `Id`, `Subject`, `From`,
`ReceivedDateTime`, `HasAttachments`, `IsRead`, `WebLink`).
**Outputs (table, `--table`).** Columns `Received | From | Subject | Att | Id`.
**Exit codes.** `0`, `2` (both `--folder` and `--folder-id`; invalid path),
`3`, `4`, `5` (`UPSTREAM_FOLDER_NOT_FOUND` on bad path / id), `6`.
**Source.** Refined request §5.3 + folder extension refined-request-folders §5.5.

---

## FR-004 — `outlook-cli get-mail <id>`

**Purpose.** Retrieve the full resource for a specific message, plus a computed
`Attachments[]` metadata summary.

**Inputs.** `<id>` positional (required), `--body <html|text|none>` (default `text`).
**REST targets.** `GET /api/v2.0/me/messages/{id}` and
`GET /api/v2.0/me/messages/{id}/attachments` (metadata only).
**Outputs (JSON).** Full message resource augmented with
`Attachments: [{ Id, Name, ContentType, Size, IsInline }, ...]`.
**Exit codes.** `0`, `2` (missing id), `3`, `4`, `5`, `6`.
**Source.** Refined request §5.4.

---

## FR-005 — `outlook-cli download-attachments <id>`

**Purpose.** Save every non-inline `FileAttachment` of a given message to a local
directory. `ItemAttachment` and `ReferenceAttachment` entries are skipped with a
recorded reason.

**Inputs.** `<id>` positional (required), `--out <dir>` (required — no default,
`ConfigurationError` if missing), `--overwrite` (default `false`), `--include-inline`
(default `false`).
**REST targets.** `GET /api/v2.0/me/messages/{id}/attachments` (list) + per-item
`GET /api/v2.0/me/messages/{id}/attachments/{attId}` (detail for `ContentBytes`).
**Outputs (JSON).** `{ messageId, outDir, saved: [{ id, name, path, size }],
skipped: [{ id, name, reason, ... }] }`. Skip reasons: `inline`, `item-attachment`,
`reference-attachment`, `content-bytes-null`, `access-denied`, `not-found`,
`unknown-attachment-type`.
**Exit codes.** `0`, `2` (missing id), `3` (missing `--out`), `4`, `5`, `6` (collision
without `--overwrite`, write error).
**Source.** Refined request §5.5; attachment semantics from
`docs/research/outlook-v2-attachments.md`.

---

## FR-006 — `outlook-cli list-calendar`

**Purpose.** List events in a time window, ordered by start time ascending.

**Inputs.** `--from <ISO8601>` (default `now`, env `OUTLOOK_CLI_CAL_FROM`),
`--to <ISO8601>` (default `now + 7d`, env `OUTLOOK_CLI_CAL_TO`), `--tz <IANA tz>`
(default system timezone, env `OUTLOOK_CLI_TZ`).
**REST target.** `GET /api/v2.0/me/calendarview?startDateTime=...&endDateTime=...&$orderby=Start/DateTime asc&$select=Id,Subject,Start,End,Organizer,Location,IsAllDay`.
**Outputs (JSON).** Array of event summaries.
**Outputs (table, `--table`).** Columns `Start | End | Subject | Organizer | Location | Id`.
**Exit codes.** `0`, `3`, `4`, `5`, `6`.
**Source.** Refined request §5.6.

---

## FR-007 — `outlook-cli get-event <id>`

**Purpose.** Retrieve the full resource for a specific calendar event.

**Inputs.** `<id>` positional (required), `--body <html|text|none>` (default `text`).
**REST target.** `GET /api/v2.0/me/events/{id}`.
**Outputs (JSON).** Full event resource (`Subject`, `Start`, `End`, `Organizer`,
`Attendees`, `Location`, `Body`, etc.).
**Exit codes.** `0`, `2` (missing id), `3`, `4`, `5`, `6`.
**Source.** Refined request §5.7.

---

## FR-008 — `outlook-cli list-folders`

**Purpose.** Enumerate mail folders under a specified parent, either direct children only
or the full sub-tree when `--recursive` is set. Part of the folder-management extension.

**Inputs.** `--parent <name-or-path-or-id>` (default `MsgFolderRoot`; accepts well-known
alias, display-name path, or `id:AAMkAGI...`), `--recursive` (default `false`),
`--include-hidden` (default `false`), `--top <N>` (default `100`, range `1..250` —
`$top` hint per `/childfolders` page; actual paging handled by the HTTP layer).
**REST targets.** `GET /api/v2.0/me/MailFolders?$top=N&$select=...` for the root level;
`GET /api/v2.0/me/MailFolders/{parentId}/childfolders?$top=N&$select=...` for children.
Pagination follows `@odata.nextLink` up to a 50-page cap per collection; whole-tree
walk bounded at 5000 visited nodes.
**Default `$select`.** `Id,DisplayName,ParentFolderId,ChildFolderCount,UnreadItemCount,
TotalItemCount,WellKnownName,CreatedDateTime,IsHidden`.
**Outputs (JSON).** Array of `FolderSummary`: `{ Id, DisplayName, ParentFolderId,
ChildFolderCount, UnreadItemCount, TotalItemCount, WellKnownName?, Path }`. `Path` is
materialized client-side from parent down using `/` separators (with `\/` and `\\`
escapes inside segments).
**Outputs (table, `--table`).** Columns `Path | Unread | Total | Children | Id` (no
truncation on `Id`).
**Exit codes.** `0`, `2` (bad flag value / invalid parent), `3`, `4`, `5`
(`UPSTREAM_FOLDER_NOT_FOUND` on bad parent; `UPSTREAM_PAGINATION_LIMIT` on cap
exceeded), `6`.
**Source.** Refined-request-folders §5.1.

---

## FR-009 — `outlook-cli find-folder <query>`

**Purpose.** Resolve a query (well-known alias, display-name path, or `id:...`) to a
single folder, returning its full metadata plus a `ResolvedVia` discriminator that
indicates how the lookup was satisfied.

**Inputs.** `<query>` positional (required — missing → `UsageError BAD_USAGE`),
`--parent <name-or-path-or-id>` (default `MsgFolderRoot`; used only when `<query>` is a
bare name), `--case-sensitive` (default `false`; default matching is Unicode NFC +
simple case-fold), `--include-hidden` (default `false`), `--first-match` (default
`false`; picks first candidate by `CreatedDateTime asc, Id asc` on ambiguity).
**REST target.** Either `GET /api/v2.0/me/MailFolders/{id}` (for `id:` form or
well-known alias) or repeated `GET /api/v2.0/me/MailFolders/{parentId}/childfolders`
(path walk, client-side match).
**Outputs (JSON).** One `ResolvedFolder`: `FolderSummary & { Path, ResolvedVia:
"wellknown" | "path" | "id" }`.
**Outputs (table, `--table`).** Two-line key/value fallback — current JSON is emitted
when no `ColumnSpec` is supplied (unchanged behaviour).
**Exit codes.** `0`, `2` (missing query; ambiguous match without `--first-match`;
invalid path — `FOLDER_PATH_INVALID`; depth cap), `3`, `4`, `5`
(`UPSTREAM_FOLDER_NOT_FOUND`), `6`.
**Source.** Refined-request-folders §5.2.

---

## FR-010 — `outlook-cli create-folder <path>`

**Purpose.** Create a new folder by path (`A/B/C`), optionally creating intermediate
parents, with idempotent-on-collision semantics.

**Inputs.** `<path>` positional (required), `--parent <name-or-path-or-id>` (default
`MsgFolderRoot`; anchor for `<path>`), `--create-parents` (default `false`;
intermediate missing parents → `UsageError FOLDER_MISSING_PARENT` without this flag),
`--idempotent` (default `false`; on collision with the existing leaf, exit 0 and
return the existing folder's id), `--display-name <name>` (default: last path segment
after unescape).
**REST target.** `POST /api/v2.0/me/MailFolders/{parentId}/childfolders` body
`{"DisplayName": "<segment>"}` for each segment that must be created. On 400 OR 409
with `error.code == "ErrorFolderExists"`, behaviour is governed by `--idempotent`.
**Outputs (JSON).** `{ created: [{ Id, DisplayName, Path, ParentFolderId, PreExisting
}, ...], leaf: { Id, Path, DisplayName, PreExisting }, idempotent: boolean }`.
**Outputs (table, `--table`).** Applied to `created`: columns `Path | Id | PreExisting`.
**Exit codes.** `0`, `2` (missing path, empty segments, path length >16,
`FOLDER_MISSING_PARENT`, attempt to create a well-known-alias-named folder at root),
`3`, `4`, `5` (other upstream 4xx/5xx), `6` (`FOLDER_ALREADY_EXISTS` on collision
without `--idempotent`).
**Source.** Refined-request-folders §5.3.

---

## FR-011 — `outlook-cli move-mail [<id>]`

**Purpose.** Move one or more messages to a destination folder. The move endpoint
creates a copy in the destination and deletes the source; the tool explicitly surfaces
the `{ sourceId, newId }` mapping per moved message because the new id differs from
the source id on v2.0.

**Inputs.**
- `<id>` positional (single message id) OR `--ids-from <path-or-->` (read one id per
  line, blank lines and `#` comments ignored; `-` reads stdin). Mutually exclusive;
  both absent or both present → `UsageError BAD_USAGE`.
- `--to <name-or-path>` OR `--to-id <folderId>`. Mutually exclusive; both absent or
  both present → `UsageError BAD_USAGE`. `--to` aliases always pre-resolve to a raw
  id via `GET /MailFolders/{alias}` before the move (consistent behaviour across
  tenants).
- `--to-parent <name-or-path-or-id>` (default `MsgFolderRoot`; anchor when `--to` is a
  bare name).
- `--continue-on-error` (default `false`; absorbs per-id failures into `failed[]`,
  still exits 5 at the end if any failed).
- `--stop-at <N>` (default `1000`, range `1..10000`; safety valve when `--ids-from`
  has more than N ids — exceeded → `UsageError BAD_USAGE` exit 2).
- `--first-match` (default `false`; disambiguation for path resolution).

**REST target.** `POST /api/v2.0/me/messages/{id}/move` body
`{"DestinationId": "<resolvedRawId>"}`. Response body is the moved message with a
**new Id**. Multi-message moves are N sequential REST calls (NG4: no `$batch`).
**Outputs (JSON).** `{ destination: { Id, Path, DisplayName }, moved: [{ sourceId,
newId }, ...], failed: [{ sourceId, error: { code, httpStatus } }, ...], summary: {
requested, moved, failed } }`.
**Outputs (table, `--table`).** Columns `Source Id | New Id | Status | Error`.
**Exit codes.** `0` (all moved), `2` (usage), `3`, `4`, `5` (any upstream failure not
absorbed by `--continue-on-error`; partial failure with `--continue-on-error` still
exits 5 while surfacing the full `failed[]` in the JSON payload), `6` (unused).
**Source.** Refined-request-folders §5.4.

---

## Cross-cutting features

### FF-001 — Global output modes

Every command supports `--json` (default) and `--table` (mutually exclusive). Progress
messages go to stderr; JSON / table output goes to stdout. `--quiet` suppresses
non-result logs.

### FF-002 — Mandatory configuration with no fallback

`OUTLOOK_CLI_HTTP_TIMEOUT_MS`, `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`, and
`OUTLOOK_CLI_CHROME_CHANNEL` are mandatory. If unset AND not passed via the
corresponding CLI flag, a typed `ConfigurationError` is raised and the process exits
with code `3`. No silent defaults. Source: refined request §8.

### FF-003 — Transparent re-authentication on 401 / expired token

On `401 Unauthorized` or when `now + 60s >= bearer.expiresAt`, the tool opens the
Playwright browser exactly once, captures a fresh Bearer + cookie jar, rewrites the
session file atomically, and retries the original REST call once. A second failure
exits `4`. `--no-auto-reauth` disables this and exits `4` on the first failure.
Source: refined request §6.4.

### FF-004 — Secret hygiene

The Bearer token and cookie values are never included in `Error.message`,
`Error.stack`, JSON output, or the optional `--log-file`. The session file is mode
`0600`; the session directory is `0700`. Source: refined request §7 and §9 (AC-PERMS,
AC-NO-SECRET-LEAK).

### FF-005 — Single-browser concurrency guard

An advisory PID lock at `<session-dir>/.browser.lock` prevents two concurrent
`outlook-cli login` (or automatic re-auth) invocations from launching two Chromes
against the same profile. Stale locks (PID no longer live, or age exceeds the login
timeout) are recovered automatically. Source: investigation §4.5.

### FF-006 — Canonical folder resolver

Every folder-aware command (`list-folders`, `find-folder`, `create-folder`,
`move-mail`, and the non-alias branch of `list-mail`) routes through a single
resolver module at `src/folders/resolver.ts`. The resolver owns: path parsing
(`/`-separated, `\/` and `\\` escape rules, 16-segment depth cap, NFC normalization),
well-known alias recognition (PascalCase match: `Inbox`, `SentItems`, `Drafts`,
`DeletedItems`, `Archive`, `JunkEmail`, `Outbox`, `MsgFolderRoot`,
`RecoverableItemsDeletions`), Unicode simple case-folding comparison (with
`--case-sensitive` override), ambiguity detection (two siblings same DisplayName
→ `UsageError FOLDER_AMBIGUOUS` exit 2 unless `--first-match`, which sorts by
`CreatedDateTime asc, Id asc`), well-known precedence at the root (well-known
alias wins over any user-created folder of the same name), and idempotent-on-
collision classification (HTTP 400 OR 409 + `error.code == "ErrorFolderExists"`).
Source: refined-request-folders §6; investigation-folders §4.2;
outlook-v2-folder-duplicate-error §4.

### FF-007 — Paginated list helper with safety caps

The HTTP client exposes `listAll<T>(path, query?, opts?)` — it issues the first
`GET`, then follows every `@odata.nextLink` **verbatim** (no query re-application)
until no more `nextLink` is emitted. Default `$top` hint: `250`. Per-collection
cap: `50` pages → `UpstreamError UPSTREAM_PAGINATION_LIMIT` exit 5 with an
actionable hint. Whole-tree walks (`list-folders --recursive`) are additionally
bounded at `5000` visited nodes. Off-host `nextLink` (any hostname other than
`outlook.office.com`) is rejected as `UpstreamError UPSTREAM_PAGINATION_OFF_HOST`.
The 401-retry-once envelope applies transparently per page. Source:
outlook-v2-folder-pagination-filter §5 + §6.
