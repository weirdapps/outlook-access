# Codebase Scan: Folder Management (addendum)

Scan date: 2026-04-21
Target request: `docs/design/refined-request-folders.md`
Base scan (must be read first): `docs/reference/codebase-scan-outlook-cli.md`

This document is **supplementary**. It assumes the base scan is accurate for
everything not touched by the folder work (tsconfig, package layout,
`.playwright-profile/`, session/auth internals, testing layout). Only the
exact plug-in points for `list-folders`, `find-folder`, `create-folder`,
`move-mail`, and the `list-mail` extension are re-documented here.

---

## 1. Command-registration pattern (where new subcommands plug in)

**File:** `src/cli.ts`
**Top-level symbol:** `main(argv)` (line 433) — builds a single `commander.Command`, registers global flags, then one `.command(...).option(...).action(...)` block per subcommand.
**Key helpers inside the file:**

- `CommandDeps` interface (line 66) — the closure each handler receives.
- `buildDeps(globalFlags)` (line 87) — resolves config, builds `sessionPath`, wires `doAuthCapture` and `createClient(session)`. Every new folder command re-uses this untouched.
- `globalOptsToFlags(g)` (line 163) — commander raw opts → `CliFlags`. Only touches global flags; no per-command concerns. **Leave as-is**.
- `resolveOutputMode(g)` (line 149) — decides `'json' | 'table'` per call.
- `emitResult(data, mode, columns?)` (line 271) — the single stdout emitter. `columns` is passed only when the payload is a flat array that has a declared `ColumnSpec`.
- `makeAction(program, handler)` (line 415) — wraps every action so it (1) builds deps, (2) calls the handler, (3) routes thrown errors to `reportError` → `process.exitCode`. **New folder commands MUST go through `makeAction`**.
- `formatErrorJson(err)` (line 297) and `exitCodeFor(err)` (line 359) — discriminate on `instanceof ConfigurationError | CliAuthError | UpstreamError | IoError | OutlookCliError | CommanderLikeError` to produce `{error: {code, message, ...}}` JSON and exit codes 3/4/5/6/2. **Folder code must raise these same classes; no new top-level class is required unless §8 of the refined spec mandates a `CollisionError`** — if added, `formatErrorJson`/`exitCodeFor` must be extended.
- `parseIntArg(v)` (line 609) — the shared integer option parser for things like `--top`, `--stop-at`.
- `LIST_MAIL_COLUMNS` (line 199) — the existing message table schema. The `Id` column must NOT have a `maxWidth` (documented rationale: IDs are truncated-ellipsis-poison for copy-paste). The same rule applies to any new `Id` / `newId` column in `move-mail` and `list-folders` tables.

**Exact anchor for existing `list-mail` registration:** `src/cli.ts:486-507`. New folder subcommands follow the identical shape (command name, description, per-command `.option(...)` chain, `.action(makeAction<Opts, PositionalTuple>(program, async (deps, g, cmdOpts, ...positional) => { const result = await <cmd>.run(deps, cmdOpts, ...positional); emitResult(result, resolveOutputMode(g), <columns?>); }))`).

**What folder work needs from it:** **extend** — add five new `.command(...)` blocks (`list-folders`, `find-folder`, `create-folder`, `move-mail`) plus two new options on `list-mail` (`--folder-id`, `--folder-parent`). Two new `ColumnSpec` constants (`LIST_FOLDERS_COLUMNS`, `MOVE_MAIL_COLUMNS`, optionally a `CREATE_FOLDER_COLUMNS`) go next to the existing ones.

---

## 2. Command handler shape (representative pattern)

### 2.1 `src/commands/list-mail.ts`

- **Symbols to read:** `ListMailDeps` (line 22), `ListMailOptions` (line 31), `ALLOWED_FOLDERS` (line 37), `DEFAULT_SELECT` (line 45), `UsageError` (line 52), `run(deps, opts)` (line 57), and the **two cross-command exports** `ensureSession(deps)` (line 115) and `mapHttpError(err)` (line 142).
- **Handler shape (canonical for the project):**
  1. Validate/resolve options against `deps.config` (falls back to `CliConfig` defaults, never to a hard-coded value).
  2. `ensureSession(deps)` → `SessionFile` (honors `--no-auto-reauth`; triggers `doAuthCapture` + `saveSession` otherwise).
  3. `deps.createClient(session)` → `OutlookClient`.
  4. Build a path that starts with `/` (enforced by the client) and a `query` record of OData params.
  5. Wrap the single `client.get<T>(path, query)` in `try { ... } catch (err) { throw mapHttpError(err); }`.
  6. Return the typed result — never format or print here (`cli.ts` owns stdout).
- **Key constraints:**
  - `UsageError` is defined locally and exported for reuse by sibling commands. Folder commands must **re-export it from `list-mail.ts` or a dedicated `src/commands/usage-error.ts` module**; do NOT fork a second class.
  - `ensureSession` / `mapHttpError` are the canonical re-usable helpers for every command; `get-mail.ts` line 15 already re-imports them from `list-mail`.
- **What folder work needs from it:**
  - `list-mail.ts` itself — **extend** (add `--folder-id` / `--folder-parent` handling, delegate non-well-known names to the resolver, keep the existing well-known fast-path).
  - `ensureSession`, `mapHttpError`, `UsageError` — **reuse verbatim** from every new command file.

### 2.2 `src/commands/get-mail.ts`

- **Symbols:** `GetMailDeps` (line 17), `GetMailOptions` (line 28), `BodyMode` (line 26), `run(deps, id, opts)` (line 34).
- Demonstrates: positional `id` validation → `UsageError`; `Promise.all` fan-out when a command needs more than one REST call; merging two responses into a typed result; delete-on-"none" body handling.
- **Pattern for `find-folder`**: exactly this shape (positional `<query>`, one REST call, return typed object).
- **What folder work needs from it:** **reuse the import pattern** (`import { ensureSession, mapHttpError, UsageError } from './list-mail';`).

### 2.3 `src/commands/download-attachments.ts`

- **Symbols:** `DownloadAttachmentsDeps` (line), `DownloadAttachmentsOptions`, `DownloadAttachmentsResult`, `run(deps, id, opts)` (lines 88-221).
- Demonstrates: multi-step REST flow (list → per-item GET), collision exit 6 via `IoError`, skip records with discriminated `reason`, two accumulator arrays in the result (`saved[]`, `skipped[]`).
- **Pattern for `move-mail`**: direct blueprint for the `moved[] / failed[] / summary` accumulators and the partial-failure-raises-exit-5 rule. The `--continue-on-error` switch mirrors how `download-attachments` records skips without aborting the whole run.
- **What folder work needs from it:** **mirror** the list + accumulator style; **reuse** `IoError` for any folder collision branch (spec §10 `FOLDER_ALREADY_EXISTS`) unless a dedicated `CollisionError` is introduced.

---

## 3. HTTP client (what folder methods must call through)

**File:** `src/http/outlook-client.ts`
**Public surface:** `OutlookClient` interface (line 31) — currently a single method:

```
get<T>(path: string, query?: Record<string, QueryValue>): Promise<T>
```

- **`createOutlookClient(opts)`** (line 66) — validates `session`, `httpTimeoutMs`, `onReauthNeeded`; keeps a _mutable_ `session` so a post-reauth call uses the new token for subsequent requests. The 401-retry-once flow lives in `doGet` (line 80): on 401 it drains the body, awaits `opts.onReauthNeeded()`, reassigns `session`, and retries exactly once. Second 401 → `throwForResponse(..., 'AFTER_RETRY')`.
- **Private helpers folder work MUST reuse (do not fork):**
  - `buildUrl(path, query)` (line 123) — `path` must start with `/`; `query` record is URL-encoded via `URLSearchParams`; OData `$`-keys pass through verbatim. Folder/mover code must build its URLs via this helper (indirectly, by calling `get` / `post`).
  - `buildHeaders(session)` (line 143) — authorative Authorization / X-AnchorMailbox / Cookie wiring. `Accept: application/json` is the only Accept. **Never construct headers elsewhere**.
  - `executeFetch(url, session, timeoutMs)` (line 204) — the single `fetch()` call site with `AbortSignal.timeout(timeoutMs)`. Any `post` helper MUST call this same path so the retry/auth envelope is unified.
  - `handleSuccessOrThrow(response, url)` (line 261) — 2xx JSON parse, empty body → `null`, non-2xx → `throwForResponse` (which distinguishes 401 paths via the `authReason` argument).
  - `mapFetchException` (for `NetworkError` / `AbortError` normalization).

**What folder work needs from it:** **extend** in place:

1. Add a `post<TBody, TRes>(path, body, query?)` method to `OutlookClient` and `createOutlookClient`, factored so the body of `doGet` is generalized to `doRequest(method, path, body?, query?)`. The 401-retry-once envelope MUST be shared; do not duplicate.
2. Add a `listAll<T>(path, query?)` helper (or a `getPaged<T>`) that walks `@odata.nextLink` up to the spec §7 cap of 50 pages, raising `UpstreamError` `UPSTREAM_PAGINATION_LIMIT` beyond that. The cap lives here, not in the resolver.
3. Reuse `buildUrl`, `buildHeaders`, `executeFetch`, `handleSuccessOrThrow`, `throwForResponse`, `mapFetchException` unchanged. **No new fetch call sites**.

No change is needed to `createOutlookClient`'s validation block, session-mutation pattern, or `onReauthNeeded` semantics — those carry folder calls correctly.

---

## 4. HTTP error taxonomy

**File:** `src/http/errors.ts`
**Shape:** the HTTP layer does NOT use a discriminated union; it uses three concrete classes:

- `OutlookHttpError` (abstract base, line 21) — `code`, `httpStatus`, `url`, `requestId`. Every message and URL is run through `redactString`.
- `AuthError` (line 52) — `reason: 'NO_AUTO_REAUTH' | 'AFTER_RETRY'`. Mapped to CLI `AuthError` in `mapHttpError` via the `reason` discriminator.
- `ApiError` (line 82) — catch-all for non-401 4xx / 5xx. `code` is set by the caller (via `codeForStatus(status)` at line 136: `403→FORBIDDEN`, `404→NOT_FOUND`, `409→CONFLICT`, `429→RATE_LIMITED`, `5xx→SERVER_ERROR`, other 4xx→`API_ERROR`).
- `NetworkError` (line 106) — pre-response failures (`fetch` TypeError, abort/timeout, socket reset). Does **not** extend `OutlookHttpError` (no httpStatus).

`codeForStatus` (line 136) — already returns `CONFLICT` for 409, which the resolver needs for the "folder already exists" path. `truncateAndRedactBody` (line 151) — the single helper for embedding upstream body snippets.

**What folder work needs from it:** **reuse verbatim** for raw HTTP. Folder-specific semantic codes (`UPSTREAM_FOLDER_NOT_FOUND`, `UPSTREAM_FOLDER_AMBIGUOUS`, `UPSTREAM_PAGINATION_LIMIT`, `FOLDER_ALREADY_EXISTS`) are **CLI-layer** codes raised by `src/config/errors.ts` `UpstreamError` and the new resolver — they are NOT added to `ApiError`. The resolver translates `ApiError{code: NOT_FOUND}` → `UpstreamError{code: UPSTREAM_FOLDER_NOT_FOUND}`, and `ApiError{code: CONFLICT}` → `UpstreamError{code: FOLDER_ALREADY_EXISTS}` (or the new `CollisionError`, per open question §13).

---

## 5. CLI-layer error taxonomy

**File:** `src/config/errors.ts`
**Classes (all subclasses of abstract `OutlookCliError`, line 14):**

| Class                                                                       | Exit | Notable fields                              | Used by                                            |
| --------------------------------------------------------------------------- | ---- | ------------------------------------------- | -------------------------------------------------- | -------------------- | --------------- | ------------------- |
| `ConfigurationError` (line 36)                                              | 3    | `missingSetting`, `checkedSources`          | mandatory env/flag miss                            |
| `AuthError` (line 67)                                                       | 4    | `code: AUTH_LOGIN_CANCELLED                 | AUTH_LOGIN_TIMEOUT                                 | AUTH_401_AFTER_RETRY | AUTH_NO_REAUTH` | capture / 401 paths |
| `UpstreamError` (line 95)                                                   | 5    | `code`, `httpStatus?`, `requestId?`, `url?` | non-401 HTTP, network, timeout                     |
| `IoError` (line 127)                                                        | 6    | `code`, `path?`                             | filesystem errors (currently the only exit-6 path) |
| `UsageError` (in `src/commands/list-mail.ts:52`, extends `OutlookCliError`) | 2    | `code: BAD_USAGE`                           | per-command argv validation                        |

`OutlookCliError.exitCode` is abstract — every new error class MUST declare it, and `cli.ts` `exitCodeFor` already falls back to `err.exitCode` for any subclass (line 365).

**What folder work needs from it:** **extend the `code` vocabularies** on existing classes (no new class unless open question §13 chooses `CollisionError`):

- `UsageError.code` — add `FOLDER_AMBIGUOUS`, `FOLDER_MISSING_PARENT`, `FOLDER_PATH_INVALID`. Because `UsageError.code` is a plain `string`, no type-level change is required; just raise with the new code string.
- `UpstreamError.code` — add `UPSTREAM_FOLDER_NOT_FOUND`, `UPSTREAM_FOLDER_AMBIGUOUS`, `UPSTREAM_PAGINATION_LIMIT` (and optionally `FOLDER_ALREADY_EXISTS` if routed as `UpstreamError` + special-case in `cli.ts`).
- `AuthError.code` — **no change** (`AUTH_401_AFTER_RETRY` / `AUTH_NO_REAUTH` already cover the folder 401 paths).
- `IoError` — **no change**. If `FOLDER_ALREADY_EXISTS` is routed via a dedicated class, introduce `CollisionError extends OutlookCliError` with `exitCode = 6` and extend `cli.ts` `formatErrorJson`/`exitCodeFor` accordingly.

**`UsageError` class location note:** the canonical `UsageError` is in `src/commands/list-mail.ts:52`, re-exported via `import { UsageError } from './list-mail';` from `get-mail.ts`. The folder work may keep this pattern, but a cleaner move is to hoist it to `src/commands/usage-error.ts` — this is the only place in the current tree where a command file exports a shared error class. Plan phase decides.

---

## 6. Domain types

**File:** `src/http/types.ts`
**Conventions (verified from existing types):**

- **REST v2.0 naming** — all fields use Outlook's PascalCase (`Id`, `DisplayName`, `ParentFolderId`, `ReceivedDateTime`, `Subject`, `From`). No re-naming to camelCase on the wire side.
- **Required vs optional** — required fields are non-optional (`Id: string`, `Subject: string`); anything Outlook documents as "may be missing" is optional with `?`. Example: `MessageSummary` (line 38) has `Subject: string` required but `From?: Recipient` optional.
- **OData envelope** — `ODataListResponse<T>` (line 181) is the canonical list wrapper: `{ '@odata.context'?, '@odata.nextLink'?, value: T[] }`.
- **No discriminated unions on the wire side** — attachments are the one exception: `Attachment = FileAttachment | ItemAttachment | ReferenceAttachment` discriminated on `@odata.type`. This is the pattern for `FolderSummary` vs (if ever needed) search-folder subtypes.
- **No `any`**. Response parsing returns `T` via `client.get<T>(...)`; the client performs no schema validation.

**What folder work needs from it:** **add a new file `src/folders/types.ts`** (per refined spec §8) for `FolderSummary`, `FolderSpec`, `ResolvedFolder`, `CreateResult`, `MoveResult`. Put the REST-shaped interfaces (`FolderSummary`, `FolderCreateRequest`) next to existing REST types _logically_, i.e. if `src/http/types.ts` keeps REST wire types, move the wire-shaped `FolderSummary` there and keep the CLI-shaped `ResolvedFolder`/`MoveResult`/`CreateResult` in `src/folders/types.ts`. Plan phase decides the split.

**Minimum new REST type (wire shape):**

```ts
export interface FolderSummary {
  Id: string;
  DisplayName: string;
  ParentFolderId?: string;
  ChildFolderCount?: number;
  UnreadItemCount?: number;
  TotalItemCount?: number;
  WellKnownName?: string; // only populated by Outlook on well-known folders
  IsHidden?: boolean;
  // Added by client (not from REST): materialized path for --recursive output.
  Path?: string;
}
```

---

## 7. Output formatter

**File:** `src/output/formatter.ts`
**Symbols:** `OutputMode` (line 15), `ColumnSpec<T>` (line 17), `formatOutput(data, mode, columns?)` (line 35), internal `truncate` (line 93), internal `padRight` (line 107).

- **No per-type formatter registry, no union.** The formatter is a generic `<T>` renderer driven by a `ColumnSpec<T>[]`. Each command passes its column spec to `emitResult` in `cli.ts`. This is the extension point.
- **JSON mode** is `JSON.stringify(data, null, 2)` — any shape works; no code change needed for new payloads.
- **Table mode** requires a `ColumnSpec<T>[]`. New column specs are added in `cli.ts` next to `LIST_MAIL_COLUMNS` and `LIST_CALENDAR_COLUMNS`.
- **ID columns must NOT set `maxWidth`** (truncation turns copy-pasted IDs into the ellipsis char and the server returns `ErrorInvalidIdMalformed`). Rule enforced by comment at `cli.ts:226`.

**What folder work needs from it:** **reuse as-is**. Add three `ColumnSpec` constants in `cli.ts` (or in per-command modules, re-exported — pick in the plan):

- `LIST_FOLDERS_COLUMNS` (`Path | Unread | Total | Children | Id`), no `maxWidth` on `Id`.
- `CREATE_FOLDER_COLUMNS` (`Path | Id | PreExisting`), no `maxWidth` on `Id`. Applied to `result.created` array.
- `MOVE_MAIL_COLUMNS` (`Source Id | New Id | Status | Error`). No `maxWidth` on `Source Id` / `New Id`.

`find-folder` returns a single object (not an array). Per spec §5.2, the table output is two-line key/value. The formatter does **not** support that today. Options: (a) fall back to JSON silently (current `emitResult` already does this when no columns spec is supplied); (b) add a second renderer mode in `formatter.ts` — `'kv'` — and extend `OutputMode`. Plan phase decides; (a) is zero-risk and matches existing `login` / `auth-check` outputs.

---

## 8. Config & global-flag precedence (no new mandatory settings)

**File:** `src/config/config.ts`
**Symbol:** `loadConfig(cliFlags)` (line 202).

**Verified precedence (for every setting):**

1. `CliFlags` value (from commander-parsed flag).
2. Env var (names in `ENV` constant at top of file).
3. Explicit default **only for spec-approved optional settings** (`sessionFilePath`, `profileDir`, `tz`, `outputMode`, `listMailTop`, `listMailFolder`, `bodyMode`, `calFrom`, `calTo`).
4. Mandatory settings (`httpTimeoutMs`, `loginTimeoutMs`, `chromeChannel`) — never defaulted; `ConfigurationError` on miss (`resolveMandatoryInt` / `resolveMandatoryString` helpers).

**What folder work needs from it:** **no change required**. Per refined spec §9, every new flag is optional and has a default documented inline (or is explicitly unset). The rule stands: no fallback defaults for mandatory config, so folder code adds zero entries to `CliConfig` and zero entries to `ENV`. Per-command defaults (e.g. `--top 100` on `list-folders`, `--stop-at 1000` on `move-mail`) are validated inside the command handler, same pattern as `list-mail.run` line 63 (`top < 1 || top > 100`).

---

## 9. Session & auth layers (no change needed)

- `src/session/store.ts` — `loadSession`, `saveSession`, `isExpired(session)` (export names confirmed via `get_symbols_overview`). Used by `ensureSession` in `list-mail.ts`; folder commands inherit this for free.
- `src/auth/browser-capture.ts` — `captureOutlookSession`, `AuthCaptureError`, `CaptureResult`. Wired in `cli.ts buildDeps` via `doAuthCapture`. Folder work does **not** touch this.
- `src/auth/jwt.ts`, `src/auth/lock.ts` — unchanged.

**What folder work needs from it:** **reuse as-is**. Every folder command goes through the exact same `ensureSession → createClient → client.get / client.post` chain.

---

## Integration summary (checklist)

To ship folder support, the work MUST touch:

**Modified (existing files):**

1. `src/cli.ts` — register 4 new subcommands (`list-folders`, `find-folder`, `create-folder`, `move-mail`); add 2 options to `list-mail` (`--folder-id`, `--folder-parent`); add 2-3 new `ColumnSpec` constants; extend `formatErrorJson` / `exitCodeFor` iff `CollisionError` is introduced.
2. `src/commands/list-mail.ts` — widen `--folder` to accept non-well-known names via the resolver; add `--folder-id` / `--folder-parent` handling; preserve the well-known fast-path verbatim; keep `ALLOWED_FOLDERS` (use it as the "skip resolver" fast-path list).
3. `src/http/outlook-client.ts` — extend `OutlookClient` with `post<TBody, TRes>(path, body, query?)` and a paging helper (e.g. `listAll<T>` or `getPaged<T>`) that follows `@odata.nextLink` up to 50 pages. Refactor internal `doGet` → `doRequest(method, ...)` so the 401-retry-once envelope is shared. Reuse `buildUrl` / `buildHeaders` / `executeFetch` / `handleSuccessOrThrow` / `throwForResponse` / `mapFetchException` unchanged.
4. `src/http/types.ts` — add `FolderSummary` (REST wire shape) and extend exports.
5. `src/config/errors.ts` — no new class (unless §13 open question chooses `CollisionError`); new code _strings_ raised against existing `UsageError` / `UpstreamError`.
6. `CLAUDE.md` — add per-subcommand entries under the `<outlook-cli>` block (AC-CLAUDEMD-UPDATED-FOLDERS).
7. `docs/design/project-design.md`, `docs/design/project-functions.MD` — update per convention.

**New files:**

8. `src/folders/resolver.ts` — `resolveFolder`, `listChildren`, `createFolderPath`, `parseFolderPath`, `buildFolderPath`, `matchesWellKnownAlias`.
9. `src/folders/types.ts` — `FolderSpec`, `ResolvedFolder`, `CreateResult`, `MoveResult` (CLI-layer shapes; REST wire types live in `src/http/types.ts`).
10. `src/commands/list-folders.ts`
11. `src/commands/find-folder.ts`
12. `src/commands/create-folder.ts`
13. `src/commands/move-mail.ts`
14. `test_scripts/ac-folders-*.ts` — one script per acceptance criterion (AC-LISTFOLDERS-ROOT … AC-401-RETRY-FOLDERS).

**Untouched (verified):** `src/auth/*`, `src/session/*`, `src/config/config.ts`, `src/output/formatter.ts`, `src/util/*`.

---

## Patterns to preserve (hard rules harvested from the existing code)

1. **`client.get<T>(path, query?)` is the only call site for HTTP.** Every command body wraps a single `try { client.get/post } catch (err) { throw mapHttpError(err); }`. Folder commands must add `client.post` to the client, not a parallel `fetch`. Verified at `src/commands/list-mail.ts:96-104`, `src/commands/get-mail.ts:55-79`, `src/commands/download-attachments.ts:111-117`.

2. **Commands never format output and never call `process.stdout`.** They return a typed value; `cli.ts emitResult` picks `json`/`table`. Verified across all three existing command modules. Folder commands follow this or table-mode flags break silently.

3. **No new fallback defaults for mandatory config.** `resolveMandatoryInt` / `resolveMandatoryString` in `src/config/config.ts:202-222` throw `ConfigurationError` on miss; folder work does not add mandatory settings (spec §9), so this rule is auto-satisfied.

4. **`ensureSession` + `mapHttpError` + `UsageError` are shared via re-export from `src/commands/list-mail.ts`.** `get-mail.ts:15` is the template (`import { ensureSession, mapHttpError, UsageError } from './list-mail';`). All four new folder commands must import from the same place (or from a hoisted `src/commands/_shared.ts` if the plan phase decides to split).

5. **IDs never get `maxWidth` in any `ColumnSpec`.** Rationale comment at `src/cli.ts:226`. Applies to every new folder `Id` / `ParentFolderId` / `newId` / `sourceId` column.

6. **Error classes are concrete (not a discriminated union); discrimination is by `instanceof`.** `src/cli.ts:297-356` (`formatErrorJson`) and `src/cli.ts:359-385` (`exitCodeFor`) cascade through `instanceof ConfigurationError | CliAuthError | UpstreamError | IoError | AuthCaptureError | OutlookCliError | CommanderLikeError`. New error codes are new _strings_ on existing classes; new error classes are only justified when the exit-code surface genuinely changes (e.g. the open §13 `CollisionError` question).

7. **401 retry-once envelope is owned by the HTTP client, not the commands.** `createOutlookClient` mutates its `session` reference on successful re-auth; a second 401 raises `HttpAuthError{reason: 'AFTER_RETRY'}` which `mapHttpError` in `list-mail.ts` converts to `CliAuthError{code: 'AUTH_401_AFTER_RETRY'}`. Folder commands inherit this transparently — they MUST NOT catch 401 themselves.

Absolute output path: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/reference/codebase-scan-folders.md`
