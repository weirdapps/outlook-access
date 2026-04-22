# Changelog (fork: weirdapps/outlook-access)

This is the **fork** of `BikS2013/outlook-tool`. Upstream's CHANGELOG is at
<https://github.com/BikS2013/outlook-tool/blob/master/CHANGELOG.md>.

Fork-only features (not in upstream):
- `--since` / `--until` / `--all` / `--max` pagination on `list-mail`
- `download-sharepoint-link` command + `sharepoint-client`/`sharepoint-schema`
  modules for SharePoint Bearer-session capture and ReferenceAttachment fetching
- `OutlookClient.listMessagesInFolderAll` auto-paginating helper

---

## [1.2.0] — 2026-04-22 (fork)

Cherry-pick from upstream `BikS2013/outlook-tool` v1.2.0 + v1.3.0
(commit `cca2f50`). Preserves all fork-only features above.

### Added
- `get-thread <id>` command — retrieves every message in a conversation by
  message id (or `conv:<conversationId>` to skip the resolve hop). Flags:
  `--body html|text|none` (default text), `--order asc|desc` (default asc).
- `OutlookClient.listMessagesByConversation(conversationId, opts?)` — public
  API method behind `get-thread`.
- `OutlookClient.countMessagesInFolder(folderId, opts?)` — server-side count
  via `$count=true`. Returns `{count, exact}`; `exact: false` when the
  server did not return `@odata.count`.
- `list-mail --from <iso|keyword>` and `--to <iso|keyword>` — date filters
  using ISO-8601 or `now` / `now + Nd` / `now - Nd` grammar.
- `list-mail --just-count` — returns only `{count, exact}` via server-side
  `$count=true`. O(1) cost regardless of mailbox size.
- `src/util/dates.ts` `parseTimestamp(raw)` — shared timestamp parser.
- `list-calendar --from/--to` — now accepts `now - Nd` keyword variant
  (inherited from shared `parseTimestamp`).
- `MessageSummary.ConversationId?: string` — optional field for `$select`.
- `ODataListResponse['@odata.count']?: number` — envelope field for count.

### Changed
- `list-mail --top` cap raised from 100 → 1000 (default still 10).
- `list-mail` rejects combining `--since/--until` with `--from/--to`
  (mutually exclusive — use `--from/--to` for new code).
- `list-mail` rejects combining `--just-count` with `--all` (count is one
  HTTP call by design; `--all` paginates).

### Deferred (downstream callers to migrate next)
- `outlook-bridge` MCP wrappers in `~/SourceCode/communications-marketplace`:
  add `outlook_get_thread` and `outlook_count_mail` tool wrappers; switch
  date filters from `--since` to `--from` (semantically equivalent).
- `second-brain` `outlook_export.py` in `~/SourceCode/second-brain`: switch
  `--since` to `--from` (no behavioral change).

---

## [1.1.1] — 2026-04-21 (fork)

### Fixed
- `outlook-cli` Node stdout/stderr pipe truncation when piping to `head`/
  `tail` etc. Caused by `process.exit()` not draining pipes. Added
  `exitWithDrain()` helper that `process.stdout.write('', cb)` before
  exiting. (PR #2)

## [1.1.0] — 2026-04-21 (fork)

### Added
- `list-mail --since <iso>` / `--until <iso>` — ISO-8601 date filters.
- `list-mail --all` / `--max <N>` — auto-paginate via `@odata.nextLink`
  with safety cap (default 10000, max 100000).
- `download-sharepoint-link <url>` command — fetches a SharePoint URL
  using the captured SharePoint Bearer session.
- `OutlookClient.listMessagesInFolderAll(folderId, opts, maxResults)` —
  auto-paginating helper backing `--all`.
- `src/http/sharepoint-client.ts`, `src/session/sharepoint-schema.ts`,
  `src/http/filter-builder.ts` — supporting modules.
- `outlook-cli login --sharepoint-host <host>` — extends interactive login
  to capture a separate Bearer session for SharePoint.
- (PR #1)
