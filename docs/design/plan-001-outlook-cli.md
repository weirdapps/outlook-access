# Plan 001 — Outlook CLI Implementation

Plan date: 2026-04-21
Inputs consumed (in priority order):

1. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/refined-request-outlook-cli.md`
2. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/investigation-outlook-cli.md`
3. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/research/playwright-token-capture.md`
4. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/research/outlook-v2-attachments.md`
5. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/reference/codebase-scan-outlook-cli.md`
6. `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/CLAUDE.md`

---

## 1. Top-Level Approach

Implement a TypeScript/Node.js CLI named `outlook-cli` that reuses the authentication of the Outlook web client (`outlook.office.com`) by launching headed Chrome via Playwright `launchPersistentContext`, installing a context-level init script that hooks `window.fetch` + `XMLHttpRequest`, and capturing the first `Authorization: Bearer <jwt>` header via `context.exposeBinding`. The captured token and Playwright cookie jar are persisted as `~/.outlook-cli/session.json` (mode `0600`, atomic write) and replayed against `outlook.office.com/api/v2.0` using the native Node `fetch` + `AbortController`. A single automatic re-auth + retry is performed on `401`. Every mandatory configuration value (HTTP timeout, login timeout, Chrome channel) is sourced via a strict flag > env precedence chain that raises a typed `ConfigurationError` (exit 3) on missing input — no silent fallbacks. Seven read-only commands (`login`, `auth-check`, `list-mail`, `get-mail`, `download-attachments`, `list-calendar`, `get-event`) are wired with `commander`, emitting JSON by default or a human table on `--table`. `vitest` is used for unit tests; end-to-end acceptance scripts live under `test_scripts/` per project convention.

---

## 2. Phases

Phases are ordered by dependency, not by chronology. Where two phases can execute concurrently, the **Can be parallelized with** field lists the sibling phase IDs.

### Phase A — Project scaffolding, deps, tsconfig, bin entry

- **Goal**: Prepare the TypeScript build surface so every later phase compiles and the CLI binary is discoverable.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/package.json` (add `bin`, `scripts.build`, `scripts.dev`, `scripts.test`, install new deps)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/tsconfig.json` (widen `include` to `["src/**/*.ts", "test_scripts/**/*.ts"]`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/cli.ts` (stub with `#!/usr/bin/env node` + commander skeleton — body filled in Phase G)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/.gitignore` (exclude `node_modules/`, `dist/`, `.playwright-profile/`, `.playwright/`, `.playwright-cli/`, `outlook_report.json`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/Issues - Pending Items.md` (created, empty sections)
- **Dependencies**: none
- **Can be parallelized with**: — (must run first)
- **Verification**:
  - `npm install` completes.
  - `npx tsc --noEmit` passes with zero errors against the stub.
  - `node -e "require('./package.json').bin['outlook-cli']"` prints `dist/cli.js`.
- **Acceptance criteria covered**: (scaffolding precursor for all; no ACs directly satisfied yet)

### Phase B — Config module (mandatory-config loader + `ConfigurationError`)

- **Goal**: Resolve every configuration value via the strict flag > env precedence chain, raising `ConfigurationError` when a mandatory setting is missing.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/config/config.ts` (loader exposing `loadConfig(argvFlags)` returning a typed `Config` object)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/config/errors.ts` (`ConfigurationError`, `AuthError`, `UpstreamError`, `IoError` classes, each with `code`, `exitCode`, `cause?`, sanitized `message`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/config.spec.ts`
- **Dependencies**: Phase A
- **Can be parallelized with**: Phase C, Phase E (types only — no runtime collisions)
- **Verification**:
  - `npm test -- config` passes (unit tests cover: mandatory missing → `ConfigurationError` with expected `missingKey`; flag beats env; env beats nothing).
  - `npx tsc --noEmit` passes.
- **Acceptance criteria covered**: AC-CONFIG-MISSING

### Phase C — Session storage module (schema, atomic write, permission bits)

- **Goal**: Persist and read the session file with schema validation, atomic `write + rename`, mode `0600` on file / `0700` on dir.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/session/schema.ts` (TypeScript types matching spec §7.2 + a `validateSessionJson(raw)` function that rejects missing / mismatched fields)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/session/store.ts` (`readSession(path)`, `writeSession(path, session)`, `deleteSession(path)`; `writeSession` uses `open(wx, 0o600) + rename`; parent dir `mkdir(..., { mode: 0o700 })` + defensive `chmod 0o700`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/util/fs-atomic.ts` (shared `atomicWrite(path, buffer, { mode, overwrite })` used by session store + attachment download)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/session-store.spec.ts`
- **Dependencies**: Phase A
- **Can be parallelized with**: Phase B, Phase E
- **Verification**:
  - `npm test -- session-store` passes (write → read round-trip equals input; mode on file is `0o600`; mode on dir is `0o700`; corrupt JSON → `IoError`).
  - Manual: `stat -f %Mp%Lp /tmp/outlook-cli-test/session.json` prints `600`.
- **Acceptance criteria covered**: AC-PERMS, AC-SESSION-REUSE (the "unchanged `capturedAt`" half is enforced here — `readSession` never rewrites)

### Phase D — Auth module (Playwright launch, init-script capture, exposeBinding, lock file)

- **Goal**: Acquire a fresh Bearer + cookie jar via headed Chrome; persist via Phase C; enforce single-browser concurrency via a PID lock.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/auth/browser-capture.ts` (exports `INIT_SCRIPT_TEXT` and `captureFirstBearerToken(context, page, timeoutMs)` per `docs/research/playwright-token-capture.md §9`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/auth/jwt.ts` (manual base64url split, extracts `exp`, `puid`, `tid`, `aud`, `scp` claims)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/auth/lock.ts` (acquire/release advisory PID lock at `<sessionDir>/.browser.lock` — `open('wx')`, stale detection via `process.kill(pid, 0)` returning `ESRCH`, expiry on age > login-timeout)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/auth/login-flow.ts` (orchestrator: acquire lock → `launchPersistentContext({ channel, headless:false })` → `exposeBinding` → `addInitScript` → `page.goto('https://outlook.office.com/mail/')` → race against close + timeout → filter cookies for `.office.com|.outlook.office.com|login.microsoftonline.com` → build session object → release lock)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/jwt.spec.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/lock.spec.ts`
- **Dependencies**: Phases A, B, C
- **Can be parallelized with**: Phase E (different modules, no shared runtime state in tests)
- **Verification**:
  - `npm test -- jwt lock` passes.
  - `npx tsc --noEmit` passes.
  - Manual smoke: `node dist/cli.js login` opens Chrome, writes session file.
- **Acceptance criteria covered**: AC-LOGIN-OK, AC-USER-CANCEL (timeout path), AC-NO-SECRET-LEAK (error classes in this phase sanitize `message`/`stack`)

### Phase E — HTTP module (REST client, header builder, auto-retry on 401, error mapping)

- **Goal**: Replay captured auth against `outlook.office.com/api/v2.0/*` with mandatory timeout, single 401 retry, and the full error taxonomy from `investigation-outlook-cli.md §4.9`.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/http/outlook-client.ts` (exports `OutlookClient` with `get(path, opts)` method; accepts `session`, `timeoutMs`, `onReauthNeeded` callback; builds `Authorization`, `X-AnchorMailbox`, `Accept`, `Cookie` headers; wraps `AbortController`; on 401 invokes callback once, rebuilds headers from the refreshed session, retries exactly once)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/http/errors.ts` (`mapHttpResponseToError(response, body)` — returns typed `UpstreamError`/`AuthError` per exit-code table; NEVER embeds the Bearer or cookie values in the error's `message`, `cause`, or `stack`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/outlook-client.spec.ts` (uses `vitest` + mocked `fetch` — no live network)
- **Dependencies**: Phases A, B, C
- **Can be parallelized with**: Phase D
- **Verification**:
  - `npm test -- outlook-client` passes (200 → returns JSON; 401 → callback fired exactly once, second 401 → `AuthError` exit 4; 403 → `UpstreamError` exit 5; `AbortError` → `UpstreamError` exit 5; token never appears in error `.message` / `.stack`).
- **Acceptance criteria covered**: AC-401-RETRY, AC-INVALID-ID, AC-NO-SECRET-LEAK (runtime path)

### Phase F — Command modules (7 commands)

- **Goal**: Implement the seven read-only commands on top of Phases B/C/D/E.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/login.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/auth-check.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/list-mail.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/get-mail.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/download-attachments.ts` (consumes `docs/research/outlook-v2-attachments.md §6` pseudocode; delegates to `src/util/filename.ts`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/list-calendar.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/commands/get-event.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/util/filename.ts` (exports `sanitizeAttachmentName`, `deduplicateFilename` per `docs/research/outlook-v2-attachments.md §5.1`)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/filename.spec.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/download-attachments.spec.ts` (mocked client)
- **Dependencies**: Phases B, C, D, E
- **Can be parallelized with**: all sibling commands can be written concurrently since they share only the `OutlookClient` interface. Recommended split: `{login, auth-check}` together (they touch Phase D), `{list-mail, get-mail, list-calendar, get-event}` together (pure HTTP/output), `{download-attachments}` alone (filename util + atomic write).
- **Verification**:
  - `npx tsc --noEmit` passes.
  - `npm test -- commands` passes (unit tests per command with mocked client).
  - Manual smoke per command against a live session (to be run by user after Phase G).
- **Acceptance criteria covered**: AC-LISTMAIL-OK, AC-GETMAIL-OK, AC-DOWNLOAD-OK, AC-LISTCAL-OK, AC-GETEVENT-OK, AC-AUTHCHECK-OK, AC-OVERWRITE-GUARD, AC-MISSING-SESSION, AC-EXPIRED-TOKEN

### Phase G — CLI entrypoint (commander wiring, output formatting, `--json` / `--table`)

- **Goal**: Wire `src/cli.ts` with commander global flags + 7 subcommands; route errors to the right exit codes; render JSON or table per flag.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/cli.ts` (full body: commander setup, global flags, per-command action stubs that call Phase F modules, top-level try/catch mapping errors to exit codes, `--quiet` / `--log-file` wiring)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/src/output/formatter.ts` (`renderJson(value)`, `renderTable(rows, columns)` — use a hand-rolled minimal table (no extra dep) to keep dependency footprint at commander-only)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/unit/formatter.spec.ts`
- **Dependencies**: Phase F
- **Can be parallelized with**: — (sequential after F)
- **Verification**:
  - `npm run build` produces `dist/cli.js`.
  - `node dist/cli.js --help` prints usage with all 7 subcommands.
  - `node dist/cli.js auth-check --help` prints per-command usage.
  - `node dist/cli.js list-mail` without any session file + `--no-auto-reauth` exits 4 with a clear message (dry-run check of the wiring).
- **Acceptance criteria covered**: AC-MISSING-SESSION (full end-to-end), AC-EXPIRED-TOKEN (full end-to-end), global flag plumbing for all ACs

### Phase H — Acceptance-criteria test scripts, CLAUDE.md docs update, verification harness

- **Goal**: Cover every AC with a dedicated `test_scripts/` file, update `CLAUDE.md` with the `<outlook-cli>` block + one child per subcommand, and register the functional requirements.
- **Files created/modified**:
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-login-ok.ts` (manual — requires user interaction)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-authcheck-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-listmail-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-getmail-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-download-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-listcal-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-getevent-ok.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-session-reuse.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-missing-session.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-expired-token.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-401-retry.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-user-cancel.ts` (manual)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-config-missing.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-perms.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-no-secret-leak.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-invalid-id.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/test_scripts/ac-overwrite-guard.ts`
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/CLAUDE.md` (append `<outlook-cli>` root block + 7 child entries)
  - `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/project-design.md` (system-level design summary — may be written in parallel with this plan)
- **Dependencies**: Phase G
- **Can be parallelized with**: — (sequential after G; within H, individual `ac-*.ts` scripts are independent)
- **Verification**:
  - `npm test` runs the full suite (unit + non-interactive AC scripts) and passes.
  - Grep check: `grep -R "bearer\|Authorization: Bearer\|cookie.*value" /tmp/outlook-cli-test-log.log` returns zero hits on the log produced by `ac-no-secret-leak`.
  - `grep -n "<outlook-cli>" CLAUDE.md` returns one hit.
  - `grep -c "<login>\|<auth-check>\|<list-mail>\|<get-mail>\|<download-attachments>\|<list-calendar>\|<get-event>" CLAUDE.md` returns 7.
- **Acceptance criteria covered**: AC-SESSION-REUSE, AC-USER-CANCEL, AC-CLAUDEMD-UPDATED, plus end-to-end coverage for every other AC (which were unit-covered in earlier phases)

### Parallelization summary

```
A  →  B  ─┐
       C  ─┼→  D  ─┐
       E  ─┘      │
                  ▼
                  F  →  G  →  H
```

- Phase A is strictly first.
- Phases B, C, E can run in parallel after A.
- Phase D needs B and C; can run in parallel with E.
- Phase F needs B, C, D, E. Within F, the seven command files can be implemented concurrently.
- Phases G and H are sequential.

---

## 3. Proposed File Tree

```
src/
  cli.ts
  commands/
    auth-check.ts
    login.ts
    list-mail.ts
    get-mail.ts
    download-attachments.ts
    list-calendar.ts
    get-event.ts
  auth/
    browser-capture.ts        # INIT_SCRIPT_TEXT + captureFirstBearerToken
    jwt.ts                    # manual base64url decode of JWT claims
    lock.ts                   # PID advisory lock at <sessionDir>/.browser.lock
    login-flow.ts             # orchestrator: lock + launch + capture + write
  http/
    outlook-client.ts         # fetch wrapper: headers + timeout + 401 retry-once
    errors.ts                 # mapHttpResponseToError, sanitized error classes
  session/
    store.ts                  # read/write/delete atomically, 0600 file / 0700 dir
    schema.ts                 # validateSessionJson, TypeScript types
  config/
    config.ts                 # precedence resolver (flag > env > error)
    errors.ts                 # ConfigurationError, AuthError, UpstreamError, IoError
  output/
    formatter.ts              # renderJson, renderTable
  util/
    filename.ts               # sanitizeAttachmentName, deduplicateFilename
    fs-atomic.ts              # atomicWrite(path, buf, { mode, overwrite })
test_scripts/
  outlook_read_recent.ts      # existing; left as historical reference
  unit/
    config.spec.ts
    session-store.spec.ts
    jwt.spec.ts
    lock.spec.ts
    outlook-client.spec.ts
    filename.spec.ts
    download-attachments.spec.ts
    formatter.spec.ts
  ac-login-ok.ts              # manual (requires user interaction)
  ac-authcheck-ok.ts
  ac-listmail-ok.ts
  ac-getmail-ok.ts
  ac-download-ok.ts
  ac-listcal-ok.ts
  ac-getevent-ok.ts
  ac-session-reuse.ts
  ac-missing-session.ts
  ac-expired-token.ts
  ac-401-retry.ts
  ac-user-cancel.ts           # manual
  ac-config-missing.ts
  ac-perms.ts
  ac-no-secret-leak.ts
  ac-invalid-id.ts
  ac-overwrite-guard.ts
  ac-claudemd-updated.ts
```

Confirmed — no changes from the prompt skeleton except adding `auth/login-flow.ts` (orchestrator distinct from the pure capture primitive) and the `test_scripts/unit/` subfolder.

---

## 4. Dependency Additions

Runtime:

```bash
npm install --save commander
```

Dev:

```bash
npm install --save-dev vitest
```

Notes:

- `playwright@^1.59.1`, `@types/node@^25.6.0`, `ts-node@^10.9.2`, `typescript@^6.0.3`, `@playwright/test@^1.59.1` are already present (see `docs/reference/codebase-scan-outlook-cli.md §1`).
- No MSAL, no `@microsoft/microsoft-graph-client`, no `axios`, no `undici`, no `jwt-decode` (we manual-parse — 6 lines, zero dep footprint).
- No table library (`cli-table3` etc.) — the `src/output/formatter.ts` hand-roll keeps the runtime dep list at exactly one: `commander`.
- Test runner decision: **vitest** (fast, zero-config with `tsconfig.json`, works with CommonJS `"type": "commonjs"`, modern `expect`/`vi.mock` API). `@playwright/test` is already installed but unused — we keep it available for future end-to-end browser scenarios without adopting it as the unit runner.
- Build: `tsc` with existing `outDir: dist`, `target: ES2022`, `module: commonjs`. New `package.json` scripts:
  ```json
  {
    "scripts": {
      "build": "tsc",
      "dev": "ts-node src/cli.ts",
      "test": "vitest run",
      "test:watch": "vitest"
    },
    "bin": { "outlook-cli": "dist/cli.js" }
  }
  ```

---

## 5. Acceptance Criteria Mapping

| AC ID               | Phase                      | Test(s)                                                                             | Notes                                                                                         |
| ------------------- | -------------------------- | ----------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| AC-LOGIN-OK         | D, F(login), G             | `test_scripts/ac-login-ok.ts` (manual)                                              | Verifies session file mode `0o600`, `account.upn` non-empty, future `tokenExpiresAt`          |
| AC-AUTHCHECK-OK     | F(auth-check), G           | `test_scripts/ac-authcheck-ok.ts`                                                   | Asserts no browser launch; cheap `GET /me` returns 200                                        |
| AC-LISTMAIL-OK      | F(list-mail), G            | `test_scripts/ac-listmail-ok.ts`                                                    | Asserts array of 5 summaries with required fields                                             |
| AC-GETMAIL-OK       | F(get-mail), G             | `test_scripts/ac-getmail-ok.ts`                                                     | Full message + `Attachments[]` metadata only                                                  |
| AC-DOWNLOAD-OK      | F(download-attachments), G | `test_scripts/ac-download-ok.ts` + `test_scripts/unit/download-attachments.spec.ts` | Byte-for-byte size match; uses `docs/research/outlook-v2-attachments.md §6` loop              |
| AC-LISTCAL-OK       | F(list-calendar), G        | `test_scripts/ac-listcal-ok.ts`                                                     | Ordered by `Start/DateTime asc`; window respected                                             |
| AC-GETEVENT-OK      | F(get-event), G            | `test_scripts/ac-getevent-ok.ts`                                                    | Asserts `Start`, `End`, `Organizer`, `Attendees` present                                      |
| AC-SESSION-REUSE    | C, G                       | `test_scripts/ac-session-reuse.ts`                                                  | Two calls → no re-auth; `capturedAt` stable across calls                                      |
| AC-MISSING-SESSION  | F(all), G                  | `test_scripts/ac-missing-session.ts`                                                | No file → triggers login; `--no-auto-reauth` → exit 4                                         |
| AC-EXPIRED-TOKEN    | E, F, G                    | `test_scripts/ac-expired-token.ts`                                                  | Stale `expiresAt` → re-auth + retry succeeds; `--no-auto-reauth` → exit 4                     |
| AC-401-RETRY        | E                          | `test_scripts/ac-401-retry.ts` + `test_scripts/unit/outlook-client.spec.ts`         | Mock-driven: first 401 → callback once + retry; second 401 → exit 4; no second browser launch |
| AC-USER-CANCEL      | D                          | `test_scripts/ac-user-cancel.ts` (manual)                                           | Close browser → exit 4 within `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`; pre-existing session untouched  |
| AC-CONFIG-MISSING   | B                          | `test_scripts/ac-config-missing.ts` + `test_scripts/unit/config.spec.ts`            | Unset env + unset flag → exit 3 with `ConfigurationError` naming the missing key              |
| AC-PERMS            | C                          | `test_scripts/ac-perms.ts` + `test_scripts/unit/session-store.spec.ts`              | `stat` checks: file `0o600`, dir `0o700`                                                      |
| AC-NO-SECRET-LEAK   | D, E                       | `test_scripts/ac-no-secret-leak.ts`                                                 | `grep` log file for token and cookie values → zero hits                                       |
| AC-INVALID-ID       | E, F(get-mail)             | `test_scripts/ac-invalid-id.ts`                                                     | Bad id → exit 5 with upstream HTTP status in payload                                          |
| AC-OVERWRITE-GUARD  | F(download-attachments)    | `test_scripts/ac-overwrite-guard.ts`                                                | Collision without `--overwrite` → exit 6; with `--overwrite` → succeeds                       |
| AC-CLAUDEMD-UPDATED | H                          | `test_scripts/ac-claudemd-updated.ts`                                               | Grep for `<outlook-cli>` root + 7 child blocks                                                |

Every AC from `refined-request-outlook-cli.md §9` is listed.

---

## 6. Risks & Mitigations

Derived from `investigation-outlook-cli.md §5`. Each risk maps to an implementation action.

| #   | Risk                                                       | Phase with mitigation                                          | Action                                                                                                                                                                                                                                                                                                        |
| --- | ---------------------------------------------------------- | -------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| R1  | Bearer expires between `auth-check` and the real REST call | E                                                              | `OutlookClient` re-checks `bearer.expiresAt` just before every call with a 60 s grace; on `now + 60s >= expiresAt` → triggers `onReauthNeeded` before the HTTP call fires.                                                                                                                                    |
| R2  | MFA re-prompted on silent re-login                         | D                                                              | `login-flow.ts` keeps the headed browser open for up to `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`; no attempt to suppress MFA UI; `captureFirstBearerToken` is idle until user completes.                                                                                                                                |
| R3  | Outlook UI change breaks "inbox reached" sentinel          | D                                                              | `login-flow.ts` treats "first captured Bearer" as the success signal (not a DOM sentinel). `page.goto('outlook.office.com/mail/', { waitUntil: 'domcontentloaded' })` is the only navigation wait; the SPA does the rest.                                                                                     |
| R4  | Captured Bearer missing a scope (e.g. `Calendars.Read`)    | F(auth-check)                                                  | `auth-check` hits both `/me/messages?$top=1` AND `/me/calendarview?...&$top=1` and surfaces the first failure, matching spec §5.2.                                                                                                                                                                            |
| R5  | Cookie domain mismatch                                     | D                                                              | `login-flow.ts` filters `context.cookies()` to `.office.com`, `.outlook.office.com`, `.login.microsoftonline.com`; `src/http/outlook-client.ts` serializes honoring `domain`, `path`, `secure`, and `httpOnly` (all written, since HTTP client sends them all). Unit test on cookie serialization in Phase E. |
| R6  | Lock file stale after SIGKILL                              | D                                                              | `lock.ts`: on `EEXIST`, `process.kill(pid, 0)` → `ESRCH` means stale → overwrite. Also: expiry if file age > `max(OUTLOOK_CLI_LOGIN_TIMEOUT_MS, 30 min)`.                                                                                                                                                     |
| R7  | Browser closed mid-capture                                 | D                                                              | `captureFirstBearerToken` already registers `page.once('close')` + `context.once('close')` (see `docs/research/playwright-token-capture.md §9`). On close → `AuthError` exit 4; `login-flow.ts` does NOT write the session file.                                                                              |
| R8  | First API request uses service worker / XHR                | D                                                              | Init script patches both `fetch` and `XMLHttpRequest` (defense-in-depth per research §3). If no token within `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` → `AuthError` exit 4 with clear message.                                                                                                                          |
| R9  | Secrets leak via error messages / log                      | B (error classes) + E (error mapper) + H (`ac-no-secret-leak`) | `ConfigurationError`/`AuthError`/`UpstreamError`/`IoError` constructors explicitly exclude `bearer.token` and cookie values from `.message` and `.stack`. `mapHttpResponseToError` never includes the request headers. AC-NO-SECRET-LEAK grep-validates.                                                      |
| R10 | Persistent profile corrupted                               | D                                                              | `login --force` rebuilds; documented in CLAUDE.md tool `<info>` block. No self-repair attempted.                                                                                                                                                                                                              |
| R11 | v2.0 API decommissioned mid-flight                         | E, F                                                           | `mapHttpResponseToError` surfaces 404/410 with a clear "endpoint may be deprecated — consider Graph migration" hint in the error payload (no silent failure).                                                                                                                                                 |
| R12 | Large attachment `ContentBytes` null                       | F(download-attachments)                                        | Per `docs/research/outlook-v2-attachments.md §4.2`: attempt detail GET; if `ContentBytes == null`, add `{ reason: "content-bytes-null", size, hint }` to `skipped[]`. Non-fatal.                                                                                                                              |
| R13 | Path traversal via attachment `Name`                       | F(download-attachments)                                        | `src/util/filename.ts` implements `sanitizeAttachmentName` + `deduplicateFilename` per `docs/research/outlook-v2-attachments.md §5.1`, including resolved-path boundary check.                                                                                                                                |

---

## 7. Ambiguities / Decisions for User Input

Genuine blockers (need user confirmation before Phase A begins):

1. **Chrome channel mandatory-config**: Spec §8 marks `OUTLOOK_CLI_CHROME_CHANNEL` as mandatory. Open Question #2 in the refined request asks whether `chrome` should be an explicit default instead. **Defaulting to mandatory** per spec as written — if the user wants to relax this, a `project-memory` note + a spec update would be required per the "no fallback" rule.

Non-blocking defaults picked (user may override before Phase A with a one-line note):

| Decision                           | Default picked                                  | Where to override                         |
| ---------------------------------- | ----------------------------------------------- | ----------------------------------------- |
| Binary name                        | `outlook-cli`                                   | `package.json#bin` (Phase A)              |
| Test runner                        | `vitest`                                        | `package.json` + phase B onwards          |
| Table library                      | none (hand-rolled in `src/output/formatter.ts`) | `src/output/formatter.ts` (Phase G)       |
| JWT parser                         | manual base64url split (no dep)                 | `src/auth/jwt.ts` (Phase D)               |
| Log file when `--log-file` omitted | no log file written                             | `src/output/formatter.ts` + CLI (Phase G) |
| Atomic-write temp-file suffix      | `.tmp.<pid>.<rand>`                             | `src/util/fs-atomic.ts` (Phase C)         |
| Lock-file expiry                   | `max(OUTLOOK_CLI_LOGIN_TIMEOUT_MS, 30 min)`     | `src/auth/lock.ts` (Phase D)              |

---

## Summary

- 8 phases (A → H) with parallelization window {B, C, E} and intra-phase parallelism inside F.
- Single runtime dep added: `commander`. Single dev dep added: `vitest`.
- Full AC coverage across unit (Phase-local) + `test_scripts/ac-*.ts` (Phase H) — every AC from `refined-request-outlook-cli.md §9` is mapped.
- Risk register from investigation §5 translated into concrete phase-scoped actions (R1–R13).
- One genuine ambiguity flagged for user confirmation: Chrome channel mandatory-vs-default policy (OQ #2 in the refined request).

Absolute output path: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/plan-001-outlook-cli.md`
