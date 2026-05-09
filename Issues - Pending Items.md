# Issues - Pending Items

## Pending

<!-- Most critical / highest priority first. -->

### MAJOR

- **[folder-isFolderExistsError-fragile] `createFolder` relies on
  parsing the upstream body out of a truncated-and-redacted error
  message string.** File:
  `<upstream-repo>/src/http/outlook-client.ts`
  (`parseErrorBody`). `throwForResponse` only embeds the upstream
  body as a 512-char snippet inside `ApiError.message` after
  `truncateAndRedactBody` runs. The `parseErrorBody` helper then tries
  to JSON-parse a prefix of that message to recover
  `{ error: { code: 'ErrorFolderExists' } }`. For typical Outlook
  responses the JSON survives verbatim, but an unusually long message
  (or a base64-like run inside the message body) could cause
  `redactString` to mangle the JSON before the predicate sees it, and
  a `CollisionError` would degrade to a plain
  `UpstreamError('UPSTREAM_HTTP_400')` (exit 5 instead of 6).
  Recommended fix: attach the parsed body object directly to
  `ApiError` (e.g. `ApiError.body?: unknown`) at throw time, so
  `isFolderExistsError` can consume the real object instead of
  re-parsing the redacted message. Requires a small `ApiError`
  signature change.

### MINOR

- **[move-mail-missing-flags] `move-mail` command is missing three
  flags specified in `refined-request-folders.md §5.4` and
  `plan-002-folders.md §P5d`: `--ids-from <file>` (read message ids
  from a file, one per line), `--to-id <rawId>` (bypass alias/path
  resolution), and `--stop-at <n>` (cap loop early with exit 2 on
  overflow).** File:
  `<upstream-repo>/src/commands/move-mail.ts`
  (`MoveMailOptions`) and
  `<upstream-repo>/src/cli.ts`
  (`move-mail` registration). Current surface (variadic positional
  `<messageIds...>` + `--to <spec>`) covers the common case; each
  missing flag can be simulated from the shell (xargs for ids-from,
  explicit id: prefix in --to for to-id, head/tail for stop-at). Blocks
  AC-MOVE-STOPAT outright and partially weakens AC-MOVE-MANY (the
  end-state is achievable but the flag surface the spec promises is
  absent). Additive fix — no backward-compatibility risk.

- **[find-folder-flag-name] `find-folder` uses `--anchor` where the
  refined spec §5.2 and `project-design.md §10.7` table both use
  `--parent`.** File:
  `<upstream-repo>/src/cli.ts`
  (`find-folder` registration). Semantically identical (the flag is
  the anchor for path-form queries, default `MsgFolderRoot`), but the
  name is inconsistent with `list-folders`, `create-folder`, and the
  documented surface. Suggested fix: rename the flag to `--parent` and
  surface a deprecated alias for `--anchor` during a transition window.

- **[nested-create-PreExisting-accuracy] On a concurrent-create race,
  `create-folder --idempotent <nested path>` may report
  `PreExisting: false` even though the leaf pre-existed.** File:
  `<upstream-repo>/src/commands/create-folder.ts`
  (`runNestedPath`). The happy path resolves the full path up front
  via `tryResolveExistingPath` (sets `PreExisting: true` correctly).
  If that resolution fails and `ensurePath` then hits a pre-existing
  leaf via pre-list detection, the command surfaces the leaf with
  `PreExisting: false`. Accurate tracking would require propagating a
  per-segment flag from `ensurePath` upward — deferred as a future
  refactor (the top-level `idempotent` flag in the payload is still
  accurate for the common path).

### MAJOR (additional review)

- **[sec-leak] Body-snippet redaction is pattern-based, not token-equality based.**
  File: `<upstream-repo>/src/http/errors.ts` (`truncateAndRedactBody`)
  - `<upstream-repo>/src/util/redact.ts` (`redactString`).
    Design §4 says the client must replace any substring equal to
    `session.bearer.token` or any `cookie.value` with `[REDACTED]` before embedding
    upstream body text in an error message. The current `redactString` only catches
    any base64-url run >100 chars. This covers JWTs and most session cookies in
    practice, but the normative contract is stricter. To close the gap, thread the
    active session into the HTTP client and do an explicit `.replaceAll(token, ...)`
  - `.replaceAll(cookie.value, ...)` pass before redactString runs. Not fixed in
    this review because it requires a signature change to `createOutlookClient`
    options.

- **[design-drift] HTTP error hierarchy diverges from design §2.2.**
  Files: `<upstream-repo>/src/http/errors.ts`,
  `<upstream-repo>/src/commands/list-mail.ts`
  (`mapHttpError`). Design contract has a single `OutlookCliError` hierarchy
  (Configuration/Auth/Upstream/Io). The implementation introduces a parallel
  `OutlookHttpError` family (`ApiError`, `AuthError`, `NetworkError`). This works
  because every command funnels errors through `mapHttpError` before re-throwing,
  but the extra layer is fragile: any future command that forgets to wrap will
  leak an `ApiError` up to `cli.ts`, where it will be treated as "UNEXPECTED"
  (exit 1). Suggested follow-up: either delete the parallel hierarchy and throw
  `UpstreamError`/`AuthError` directly from the http layer, or add a generic
  `err instanceof OutlookHttpError` → `UpstreamError` map in `cli.ts`.

### MINOR / NIT

- **[dedup] `deduplicateFilename` uses an in-memory `Set` instead of the
  filesystem.** File:
  `<upstream-repo>/src/util/filename.ts`. Design §2.11
  specifies an async function that checks for existing files on disk and caps
  attempts at 999. The current implementation caps at 10 000 and only tracks
  names generated in the current batch. Behaviour is correct for a single
  `download-attachments` call because the `atomicWriteBuffer` call with
  `overwrite:false` surfaces on-disk collisions via `IO_WRITE_EEXIST`, but the
  API signature drift means future callers that pass only a bare name (without
  the Set) get no dedup at all.

- **[login-save-twice] `login` command saves the session file twice.** File:
  `<upstream-repo>/src/commands/login.ts` line 80 +
  `<upstream-repo>/src/cli.ts` line 99 (inside
  `doAuthCapture`). The first save happens inside the injected `doAuthCapture`;
  the second happens immediately after in `login.run`. The second call is an
  idempotent rewrite, so no harm done, but it doubles disk IO and is confusing.

- **[signature-drift] `sanitizeAttachmentName` signature.** File:
  `<upstream-repo>/src/util/filename.ts`. Design §2.11
  specifies `sanitizeAttachmentName(raw: string | null | undefined, fallback:
string)`; the implementation is `sanitizeAttachmentName(raw: string)` with a
  hard-coded fallback of `"attachment"`. All current callers pass a string, but
  the API deviates from the normative contract.

- **[race-toctou] `atomicWriteBuffer` `overwrite:false` still has a TOCTOU.**
  File: `<upstream-repo>/src/util/fs-atomic.ts`. The
  sequence `access(finalPath)` + `rename(tmp, finalPath)` matches design §2.10
  step 9 but `rename()` silently replaces an existing target on POSIX, so a file
  that appears between the check and the rename will be clobbered. The
  POSIX-idiomatic fix is `link(tmp, finalPath)` + `unlink(tmp)` which returns
  EEXIST atomically when the target exists. Low priority because only the
  attachments path hits `overwrite:false` and that directory is user-chosen.

- **[config-error-mix] `download-attachments --out` missing uses
  `ConfigurationError`.** File:
  `<upstream-repo>/src/commands/download-attachments.ts`
  line 98. The refined spec §5.5 specifies exit 3 for missing `--out`, which is
  consistent with `ConfigurationError`. Name-wise this is a command-level option
  not an "environment" setting, but exit-code-wise the behaviour is correct.

- **[redundant-check] `createOutlookClient` rejects zero/negative
  `httpTimeoutMs` with a plain `Error`.** File:
  `<upstream-repo>/src/http/outlook-client.ts` line 71.
  This can never fire because `loadConfig` already rejects non-positive timeouts
  via `ConfigurationError`. Harmless defensive code, but if it ever does fire
  it will surface as UNEXPECTED (exit 1) instead of CONFIG_MISSING (exit 3).

- **[auth-capture-errors] `AuthCaptureError` does not extend `OutlookCliError`.**
  File: `<upstream-repo>/src/auth/browser-capture.ts`
  lines 58-68. The CLI top-level handler does check for it explicitly
  (`cli.ts` lines 321, 338), so exit code 4 is emitted correctly. However this
  is another spot of hierarchy drift; consolidating into `AuthError` from
  `config/errors` would simplify the taxonomy.

## Completed

<!-- Completed items moved here. -->

### BLOCKER / MAJOR fixes applied during 2026-04-21 folder-management review (Phase 7)

- **[fixed] `list-mail` did not accept display-name paths in `--folder`
  and treated `--folder-parent` as a third mutually-exclusive flag.**
  File: `src/commands/list-mail.ts`. The mutual-exclusion rule was
  corrected to `--folder` XOR `--folder-id` (per design §10.7) and the
  value of `--folder` is now routed through the resolver when it is
  neither an original-five fast-path alias nor a direct id. Also added
  two additional validations: `--folder-parent` with `--folder-id` →
  exit 2; `--folder-parent` without `--folder` → exit 2. Fixes
  AC-LISTMAIL-PATH and preserves AC-LISTMAIL-WELLKNOWN-BACKCOMPAT.

- **[fixed] `create-folder <nested-path>` without `--idempotent` did
  not raise `CollisionError` when the leaf pre-existed (pre-list
  detection path).** File: `src/folders/resolver.ts` (`ensurePath`).
  The walk previously advanced silently on any pre-existing segment,
  which caused a nested non-idempotent re-run to return success with
  `PreExisting: false` instead of exit 6 (AC-CREATE-COLLISION).
  `ensurePath` now throws `CollisionError('FOLDER_ALREADY_EXISTS')`
  when the LEAF segment pre-exists and `idempotent === false`.
  Intermediate segments still advance without POST so
  `--create-parents` remains strictly about missing parents.

- **[fixed] `create-folder --parent <anchor>` was silently ignored
  for nested paths (always anchored at `MsgFolderRoot`).** File:
  `src/folders/resolver.ts` (`ensurePath`) +
  `src/commands/create-folder.ts` (`runNestedPath`). `ensurePath`
  now accepts an optional `anchor: FolderSpec` and resolves it via
  `resolveFolder` before the walk. `runNestedPath` parses `--parent`,
  passes it into both `ensurePath` and `tryResolveExistingPath`, and
  applies the "well-known alias leaf at root is forbidden" validation
  only when the anchor resolves to `MsgFolderRoot`.

### BLOCKER / MAJOR fixes applied during 2026-04-21 code review

- **[fixed] 401 + `--no-auto-reauth` now yields `AUTH_NO_REAUTH`, not
  `AUTH_401_AFTER_RETRY`.** Added a `reason` discriminator to `http/errors
AuthError`, threaded through `outlook-client.doGet`, and updated
  `list-mail.mapHttpError` to emit the correct CLI code. Matches design §2.8 /
  §4.

- **[fixed] `download-attachments` now uses `atomicWriteBuffer` (fsync +
  rename) with the `overwrite` flag instead of `fs.writeFile`.** Matches design
  §2.13.5 step 8 and gives us torn-write protection + a proper EEXIST guard.

- **[fixed] `atomicWriteBuffer` no longer forces the parent directory to
  mode 0o700.** Added an opt-in `parentDirMode` option. The session file still
  gets 0o700; the user's `download-attachments --out` directory now keeps its
  own mode (user umask). Matches design §2.13.5 step 1.

- **[fixed] `download-attachments` `ensureOutDir` no longer passes
  `mode: 0o700`.** Same motivation as above.
