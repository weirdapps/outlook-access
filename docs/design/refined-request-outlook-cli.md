# Refined Request: Outlook CLI Tool

## 1. Summary

Build a TypeScript/Node.js CLI tool that lets the signed-in user read their Microsoft 365 / work Outlook mailbox and calendar by reusing the authentication of the Outlook web client (`outlook.office.com`). When local auth cache is missing, expired, or rejected, the tool launches a headed Chrome browser via Playwright, lets the user log in interactively, captures both cookies and the Bearer token used by the web client (by hooking `window.fetch`), and persists them safely under a user-owned dot-folder. All data access goes through the verified `outlook.office.com/api/v2.0` REST endpoints. The tool exposes read-only commands for inbox listing, message detail, attachment download, calendar listing, and event detail. No sending, no Graph API, no fallback configuration defaults.

## 2. Goals

- G1. Provide a single TypeScript CLI binary (e.g. `outlook-cli`) that can be invoked repeatedly without re-authentication as long as the cached session is valid.
- G2. Detect missing / expired / rejected auth and transparently re-drive the interactive login through Playwright-controlled Chrome.
- G3. Capture both session cookies and the MSAL Bearer token used by the Outlook web client, by hooking `window.fetch` on the first `outlook.office.com` API call.
- G4. Persist captured secrets to a file under `$HOME`, with `0600` permissions.
- G5. Expose the following commands, each producing deterministic, scriptable output (JSON by default, human table optional): `login`, `auth-check`, `list-mail`, `get-mail`, `download-attachments`, `list-calendar`, `get-event`.
- G6. On `401 Unauthorized` from the REST API, trigger a single automatic re-auth and retry; never loop indefinitely.
- G7. Enforce the project's "no fallback defaults for mandatory config" rule by raising explicit, typed errors.
- G8. Register the tool and all its sub-commands inside `CLAUDE.md` using the `<toolName>/<objective>/<command>/<info>` XML format.

## 3. Non-Goals

- NG1. Sending, replying to, or forwarding email.
- NG2. Creating, updating, or deleting calendar events.
- NG3. Modifying any server-side state (no mark-as-read, no move, no delete, no flag).
- NG4. Using Microsoft Graph API (`graph.microsoft.com`). Only `outlook.office.com/api/v2.0` is in scope.
- NG5. Decrypting tokens from `localStorage` (MSAL AES-GCM blobs). Tokens are captured via `window.fetch` interception only.
- NG6. Supporting personal `outlook.com` accounts. Only Microsoft 365 / work accounts on `outlook.office.com`.
- NG7. Background daemon / headless automatic refresh without a visible browser. Re-auth is always user-driven in a headed window.
- NG8. macOS Keychain / Secret Service / Windows Credential Manager integration (file-based `0600` is the contract for this iteration).
- NG9. Multi-account support. One active account per session file.
- NG10. Shell completions, man pages, telemetry.

## 4. User Stories

- US1. **login** — As a user, I run `outlook-cli login` to open a Chrome window, sign into Outlook, and have the tool capture and store my cookies and Bearer token so subsequent commands work without another login.
- US2. **auth-check** — As a user, I run `outlook-cli auth-check` to verify that my cached session is present and still accepted by the Outlook API, without performing any other action.
- US3. **list-mail** — As a user, I run `outlook-cli list-mail` to see my N most recent Inbox messages (subject, from, received date, id, hasAttachments) so I can pick one to inspect.
- US4. **get-mail** — As a user, I run `outlook-cli get-mail <id>` to retrieve the full content of a message (headers, recipients, body, attachment metadata) for scripting or reading.
- US5. **download-attachments** — As a user, I run `outlook-cli download-attachments <id> --out <dir>` to save every attachment of a given message to a local folder.
- US6. **list-calendar** — As a user, I run `outlook-cli list-calendar` to see my upcoming calendar events within a configurable window.
- US7. **get-event** — As a user, I run `outlook-cli get-event <id>` to see the full details of one event (organizer, attendees, location, body, times).

## 5. CLI Surface

Global options (apply to every subcommand):

- `--json` (default `true`) — emit JSON to stdout.
- `--table` — emit a human-readable table to stdout (mutually exclusive with `--json`).
- `--quiet` — suppress non-result logs (progress messages go to stderr when set, or are silenced).
- `--timeout <ms>` — HTTP timeout per REST call. No default; if unset, the tool reads `OUTLOOK_CLI_HTTP_TIMEOUT_MS` from env. If neither is set, an explicit exception is raised.
- `--session-file <path>` — override session file path. Default: `$HOME/.outlook-cli/session.json`.
- `--no-auto-reauth` — on 401, fail instead of re-opening the browser.

Exit codes:

- `0` success
- `2` invalid arguments / usage
- `3` configuration error (missing mandatory config, no fallback allowed)
- `4` auth failure (user cancelled login, or re-auth could not capture a token)
- `5` upstream API error (non-401 HTTP error from Outlook)
- `6` IO error (e.g. cannot write session file, cannot create output dir)

### 5.1 `outlook-cli login`

- Arguments: none.
- Options:
  - `--force` — ignore existing cache and always open the browser.
- Behavior: opens headed Chrome via Playwright; waits for the user to reach the Outlook inbox; hooks `window.fetch` beforehand to capture the first Bearer token; collects cookies; writes the session file.
- Output (JSON):
  ```json
  {
    "status": "ok",
    "sessionFile": "/Users/.../.outlook-cli/session.json",
    "tokenExpiresAt": "2026-04-21T13:05:00Z",
    "account": { "puid": "...", "tenantId": "...", "upn": "user@example.com" }
  }
  ```

### 5.2 `outlook-cli auth-check`

- Arguments: none.
- Options: none.
- Behavior: loads session file; performs a cheap REST call (e.g. `GET /api/v2.0/me`); reports validity. Does NOT auto-reauth.
- Output (JSON):
  ```json
  {
    "status": "ok" | "expired" | "missing" | "rejected",
    "tokenExpiresAt": "2026-04-21T13:05:00Z" | null,
    "account": { "upn": "user@example.com" } | null
  }
  ```

### 5.3 `outlook-cli list-mail`

- Arguments: none.
- Options:
  - `-n, --top <N>` — number of messages (default `10`, range `1..100`).
  - `--folder <name>` — default `Inbox`. Only well-known folder names are accepted in this iteration (`Inbox`, `SentItems`, `Drafts`, `DeletedItems`, `Archive`).
  - `--select <comma-list>` — override the selected fields; default is the set below.
- Default selected fields: `Id, Subject, From, ReceivedDateTime, HasAttachments, IsRead, WebLink`.
- REST target: `GET /api/v2.0/me/MailFolders/<folder>/messages?$top=N&$orderby=ReceivedDateTime desc&$select=...`.
- Output (JSON): an array of message summaries matching the selected fields.
- Output (table, when `--table`): columns `Received | From | Subject | Att | Id`.

### 5.4 `outlook-cli get-mail <id>`

- Arguments:
  - `<id>` — required message id (positional). If missing, exit 2.
- Options:
  - `--body <html|text|none>` — include body as HTML, plain text, or omit (default `text`).
- REST target: `GET /api/v2.0/me/messages/{id}`.
- Output (JSON): the full message resource plus a computed `Attachments[]` summary (`Id, Name, ContentType, Size, IsInline`) obtained from `GET /api/v2.0/me/messages/{id}/attachments` (metadata only, no content bytes).

### 5.5 `outlook-cli download-attachments <id>`

- Arguments:
  - `<id>` — required message id.
- Options:
  - `--out <dir>` — required; no default. If missing, exit 3 (config error).
  - `--overwrite` — overwrite existing files (default `false`; without this flag, colliding files cause exit 6).
  - `--include-inline` — also save inline attachments (default `false`).
- REST target: `GET /api/v2.0/me/messages/{id}/attachments` then per-item `GET /api/v2.0/me/messages/{id}/attachments/{attId}` (or `$value` where applicable for file attachments).
- Output (JSON):
  ```json
  {
    "messageId": "...",
    "outDir": "/abs/path",
    "saved": [
      { "id": "...", "name": "report.pdf", "path": "/abs/path/report.pdf", "size": 12345 }
    ],
    "skipped": [
      { "id": "...", "name": "logo.png", "reason": "inline" }
    ]
  }
  ```

### 5.6 `outlook-cli list-calendar`

- Arguments: none.
- Options:
  - `--from <ISO8601>` — window start. If omitted, reads `OUTLOOK_CLI_CAL_FROM` from env. If neither set, default is "now" (this is the ONE explicit default allowed; see Configuration).
  - `--to <ISO8601>` — window end. If omitted, reads `OUTLOOK_CLI_CAL_TO`. If neither set, default is "now + 7 days".
  - `--tz <IANA tz>` — timezone override, defaults to system timezone (allowed default).
- REST target: `GET /api/v2.0/me/calendarview?startDateTime=...&endDateTime=...&$orderby=Start/DateTime asc&$select=Id,Subject,Start,End,Organizer,Location,IsAllDay`.
- Output (JSON): array of event summaries.
- Output (table): `Start | End | Subject | Organizer | Location | Id`.

### 5.7 `outlook-cli get-event <id>`

- Arguments:
  - `<id>` — required event id.
- Options:
  - `--body <html|text|none>` — include body (default `text`).
- REST target: `GET /api/v2.0/me/events/{id}`.
- Output (JSON): full event resource.

## 6. Auth Flow

### 6.1 State discovery (runs before every command except `login --force`)

1. Resolve session file path (flag > env `OUTLOOK_CLI_SESSION_FILE` > `$HOME/.outlook-cli/session.json`).
2. If the file does not exist → state = `missing`.
3. If it exists but cannot be parsed or has wrong schema → state = `corrupt`; treat as `missing` for the next step (but log a warning to stderr).
4. If parsed: check `tokenExpiresAt`. If `now + 60s >= tokenExpiresAt` → state = `expired`.
5. Otherwise → state = `present`.

### 6.2 Cached auth present

1. Build request with headers:
   - `Authorization: Bearer <bearerToken>`
   - `X-AnchorMailbox: PUID:<puid>@<tenantId>`
   - `Accept: application/json`
   - `Cookie: <serialized cookie jar for outlook.office.com>`
2. Send the REST call.
3. On `2xx` → return result.
4. On `401` → go to 6.4 (re-auth), then retry ONCE. On second `401` → exit code 4.
5. On any other non-2xx → exit code 5 with body/status in error payload.

### 6.3 Cached auth missing or expired (or `login` / `login --force`)

1. Launch Playwright with a **headed** Chrome (channel `chrome`), using a dedicated persistent profile dir under `$HOME/.outlook-cli/playwright-profile/`.
2. Before navigation, install an init script that hooks `window.fetch`:
   - For every request to `https://outlook.office.com/` (ows, api/v2.0, or Graph bridge), capture `request.headers['authorization']` and post it back to the Node side via `page.exposeBinding` / a dedicated message channel.
3. Navigate to `https://outlook.office.com/mail/`.
4. Wait for the user to complete login and land on the inbox (detect by URL and/or a DOM sentinel such as the mail list container).
5. Once the first Bearer token is captured AND the inbox is reached:
   - Read all cookies for `.office.com`, `.outlook.office.com`, `login.microsoftonline.com`.
   - Extract `puid` and `tenantId` from the JWT (or from a follow-up `GET /api/v2.0/me` call once the token is known).
   - Compute `tokenExpiresAt` from the JWT `exp` claim.
6. Write the session file atomically (`write + rename`) with mode `0600`.
7. Close the browser.

### 6.4 On 401 during a REST call

1. If `--no-auto-reauth` is set → exit 4 with a clear message.
2. Otherwise, invoke 6.3.
3. After the new session is persisted, retry the original REST call exactly once.
4. If it still returns 401 → exit 4.

### 6.5 User cancellation

- If the browser window is closed before login completes, or no token is captured within `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` (mandatory config, see §8) → exit 4 with message "login not completed".

## 7. Session Storage

### 7.1 Files and permissions

- Session dir: `$HOME/.outlook-cli/` — mode `0700`.
- Session file: `$HOME/.outlook-cli/session.json` — mode `0600`.
- Playwright persistent profile: `$HOME/.outlook-cli/playwright-profile/` — mode `0700` (contains Chrome profile data; treat as sensitive).
- Log file (optional, only when `--log-file` given): mode `0600`.

### 7.2 Session file schema (JSON)

```json
{
  "version": 1,
  "capturedAt": "2026-04-21T12:05:00Z",
  "account": {
    "upn": "user@example.com",
    "puid": "10032xxxxxxxxxxx",
    "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  },
  "bearer": {
    "token": "<JWT>",
    "expiresAt": "2026-04-21T13:05:00Z",
    "scopes": ["..."],
    "audience": "https://outlook.office.com/"
  },
  "cookies": [
    {
      "name": "X-OWA-CANARY",
      "value": "...",
      "domain": ".outlook.office.com",
      "path": "/",
      "expires": 1777777777,
      "httpOnly": true,
      "secure": true,
      "sameSite": "None"
    }
  ],
  "anchorMailbox": "PUID:10032xxxxxxxxxxx@xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

Field meanings:

- `version` — schema version; bump on breaking changes.
- `capturedAt` — ISO8601 UTC, when this file was last written.
- `account.upn` / `puid` / `tenantId` — used for building `X-AnchorMailbox` and for display.
- `bearer.token` — raw JWT; never logged.
- `bearer.expiresAt` — derived from JWT `exp`; used by `auth-check` and pre-call gating.
- `bearer.scopes` / `audience` — informational; surfaced by `auth-check`.
- `cookies[]` — full cookie jar, Playwright-shaped.
- `anchorMailbox` — pre-computed convenience.

### 7.3 Write rules

- Always write to a sibling temp file and `rename` into place to avoid partial writes.
- Set mode `0600` BEFORE writing the sensitive content (open with `O_CREAT|O_WRONLY|O_TRUNC`, `0600`).
- Never echo `bearer.token` or cookie values to stdout/stderr, even at debug log level.

## 8. Configuration

Precedence (highest wins): CLI flag > environment variable > session file (for runtime state only) > hard-coded value (only where this spec explicitly allows a default).

**Mandatory configuration (NO fallback — raise `ConfigurationError` with exit code 3 if unset):**

| Name | Env var | CLI flag | Purpose |
|---|---|---|---|
| HTTP timeout (ms) | `OUTLOOK_CLI_HTTP_TIMEOUT_MS` | `--timeout` | Per-REST-call timeout |
| Login timeout (ms) | `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` | `--login-timeout` | Max time to wait for user login + token capture |
| Playwright Chrome channel | `OUTLOOK_CLI_CHROME_CHANNEL` | `--chrome-channel` | Which Chrome to launch (`chrome`, `msedge`, etc.) |

**Optional configuration (defaults explicitly allowed by this spec):**

| Name | Env var | CLI flag | Default |
|---|---|---|---|
| Session file path | `OUTLOOK_CLI_SESSION_FILE` | `--session-file` | `$HOME/.outlook-cli/session.json` |
| Playwright profile dir | `OUTLOOK_CLI_PROFILE_DIR` | `--profile-dir` | `$HOME/.outlook-cli/playwright-profile` |
| Output mode | — | `--json` / `--table` | `--json` |
| Mail top N | — | `-n / --top` | `10` |
| Mail folder | — | `--folder` | `Inbox` |
| Calendar window start | `OUTLOOK_CLI_CAL_FROM` | `--from` | `now` |
| Calendar window end | `OUTLOOK_CLI_CAL_TO` | `--to` | `now + 7d` |
| Timezone | `OUTLOOK_CLI_TZ` | `--tz` | system timezone |
| Body format | — | `--body` | `text` |

Any mandatory value that is missing results in an immediate, typed `ConfigurationError` carrying the name of the missing setting and the precedence chain that was checked.

## 9. Acceptance Criteria

Each AC below must be covered by at least one test script under `test_scripts/`.

### Passing scenarios

1. **AC-LOGIN-OK** — Running `outlook-cli login` with no prior session opens a headed Chrome window; after the user completes login, a session file is written at the configured path with mode `0600`, and stdout contains a JSON object with `status: "ok"`, a non-empty `account.upn`, and a future `tokenExpiresAt`.
2. **AC-AUTHCHECK-OK** — With a valid session, `outlook-cli auth-check` prints `{"status":"ok", ...}` and exits 0 WITHOUT launching a browser.
3. **AC-LISTMAIL-OK** — With a valid session, `outlook-cli list-mail -n 5` returns an array of 5 message summaries, each containing `Id`, `Subject`, `From`, `ReceivedDateTime`, `HasAttachments`.
4. **AC-GETMAIL-OK** — With a valid session, `outlook-cli get-mail <validId>` returns a JSON object including `Id`, `Subject`, `Body`, `ToRecipients`, and an `Attachments` array with metadata only.
5. **AC-DOWNLOAD-OK** — `outlook-cli download-attachments <idWithAttachments> --out ./tmp-attachments` writes every non-inline attachment to disk; the returned JSON `saved[]` matches the files present on disk byte-for-byte in count and size.
6. **AC-LISTCAL-OK** — `outlook-cli list-calendar --from <ISO> --to <ISO>` returns an ordered array of events within the window.
7. **AC-GETEVENT-OK** — `outlook-cli get-event <validId>` returns the full event resource with `Start`, `End`, `Organizer`, `Attendees`.
8. **AC-SESSION-REUSE** — Two consecutive calls to `outlook-cli list-mail -n 1` using the same session file produce two successful responses, and the session file's `capturedAt` is NOT modified between them (no unnecessary re-auth).

### Failing / edge scenarios

9. **AC-MISSING-SESSION** — With no session file, `outlook-cli list-mail` automatically invokes the login flow; if `--no-auto-reauth` is passed instead, it exits with code 4 and a clear error.
10. **AC-EXPIRED-TOKEN** — With a session whose `bearer.expiresAt` is in the past, any data command transparently triggers the re-auth flow (browser opens once), then succeeds. With `--no-auto-reauth`, the same situation exits 4.
11. **AC-401-RETRY** — When the REST API returns 401 despite a non-expired cached token, the tool re-opens the browser exactly ONCE, refreshes the session, retries, and succeeds. A second 401 causes exit 4 without a second browser launch.
12. **AC-USER-CANCEL** — If the user closes the Chrome window without logging in, the tool exits 4 within `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` and leaves any pre-existing session file untouched.
13. **AC-CONFIG-MISSING** — If `OUTLOOK_CLI_HTTP_TIMEOUT_MS` is unset AND `--timeout` is not passed, the tool exits 3 with a `ConfigurationError` naming `OUTLOOK_CLI_HTTP_TIMEOUT_MS`. No silent default.
14. **AC-PERMS** — After any write of the session file, its mode is exactly `0600` and the parent dir is `0700`. Verified in test.
15. **AC-NO-SECRET-LEAK** — With `--log-file` enabled at debug level, the produced log file contains NO occurrences of `bearer.token` value and NO cookie values.
16. **AC-INVALID-ID** — `outlook-cli get-mail bogus-id` exits 5 with an error payload containing the upstream HTTP status and a redacted message.
17. **AC-OVERWRITE-GUARD** — `download-attachments` into a directory where a file of the same name already exists exits 6 unless `--overwrite` is passed.
18. **AC-CLAUDEMD-UPDATED** — `CLAUDE.md` is updated with tool docs per project conventions (one top-level `<outlook-cli>` entry plus one child entry per subcommand, each using the `<toolName>/<objective>/<command>/<info>` format).

## 10. Open Questions

1. **Binary name** — confirm `outlook-cli` as the npm `bin` name and the invocation used in CLAUDE.md `<command>` blocks. (Assumed yes unless changed.)
2. **Chrome channel default** — this spec marks `OUTLOOK_CLI_CHROME_CHANNEL` as mandatory (no fallback). Confirm that this is acceptable, vs. allowing `chrome` as an explicit default.
