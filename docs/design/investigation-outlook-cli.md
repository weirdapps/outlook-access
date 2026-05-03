# Investigation: Outlook CLI — Auth & REST Access Strategy

Investigation date: 2026-04-21
Inputs consumed:

- `docs/design/refined-request-outlook-cli.md`
- `docs/reference/codebase-scan-outlook-cli.md`
- Proof-of-concept outcome from the main chat: Playwright launched headed Chrome, an
  init-script hooked `window.fetch`, a Bearer token was captured from the first
  `outlook.office.com/api/v2.0/*` call, and that same token worked via `curl` against
  `/me/messages` and `/me/calendarview` from outside the browser. Token TTL ~1h; cookies
  persist longer.

---

## 1. Problem restatement

We are building a read-only TypeScript/Node CLI that lets the signed-in user query their
Microsoft 365 / work Outlook mailbox and calendar by piggy-backing on the authentication
of the Outlook web client (`outlook.office.com`). The spec pins the tool to the REST
surface `outlook.office.com/api/v2.0` (not Microsoft Graph) and explicitly forbids any
attempt to decrypt the MSAL tokens stored in the browser's `localStorage`.

This is nontrivial for three reasons:

1. **MSAL token storage is opaque.** Outlook web uses MSAL.js, and the access tokens it
   persists to `localStorage` are AES-GCM-encrypted using a key that lives in a cookie
   (`msal.cache.encryption`). We cannot pluck a usable Bearer out of storage; we must
   intercept the token while it is in flight.
2. **The Bearer expires in ~1 hour,** which is shorter than many user sessions. Cached
   cookies (`X-OWA-CANARY`, `SignInStateCookie`, ESTS cookies) survive much longer. The
   tool must distinguish "token expired but session valid" (silent re-login via the
   persistent Chrome profile) from "session gone" (interactive login, MFA possible).
3. **The web client is a moving target.** Header shape (`X-AnchorMailbox`,
   `prefer` tokens, `action` scopes), DOM sentinels for "inbox reached", and the exact
   URL patterns of authenticated API calls can shift between Outlook releases. We need a
   capture strategy robust enough to survive minor drift and explicit failure modes when
   it does not.

---

## 2. Approach comparison

### A. Browser-hook approach (captured Bearer via `window.fetch` hook)

**What it is:** Launch headed Chrome via Playwright, install a page init-script that
wraps `window.fetch` (and optionally `XMLHttpRequest`), intercept the first authenticated
call to `outlook.office.com/api/v2.0/*`, extract the `Authorization: Bearer <JWT>`
header, relay it to Node via `page.exposeBinding`, and harvest cookies from the browser
context. This is the POC path and the spec-locked choice.

- **Pros**
  - Requires no tenant admin consent, no app registration, no Entra changes. A bank /
    enterprise user can use it today.
  - Reuses exactly the same session the user already has in Outlook web — same scopes,
    same audience (`outlook.office.com`), same conditional-access posture.
  - POC already proved it works end-to-end (headed login -> captured Bearer -> curl
    against `/me/messages` and `/me/calendarview` succeeds).
- **Cons**
  - Token lifetime is ~1 hour; must be refreshed by re-driving the browser (usually
    silent against the persistent profile, occasionally MFA-prompted).
  - Sensitive to Outlook web UI / DOM changes: the "inbox reached" sentinel and the URL
    pattern used to filter the first authenticated fetch can drift.
  - Secrets (Bearer + cookies) end up on disk under `$HOME/.outlook-cli/`; mitigated by
    mode `0600` and atomic writes but not on par with OS keystore integration.
- **Failure modes**
  - Token expiry between the `auth-check` call and the real REST call (the 60 s
    pre-call grace window in spec §6.1 is the mitigation).
  - `Authorization` header drift (e.g. additional `prefer` / `action` tokens baked into
    the real request that we didn't replay) causing 401 / 403.
  - Outlook renames the inbox DOM sentinel and the login-complete wait loop times out.
  - First API call does not carry a Bearer (rare — the OWS endpoints use it from the
    very first call, but a navigation that hits a cached shell first could delay
    capture).
- **Effort: S** (spec-locked; POC already validates the approach).

### B. Entra ID app registration + OAuth2 PKCE (native CLI)

**What it is:** Register a public-client app in the user's Entra tenant, declare the
delegated scopes we need (`Mail.Read`, `Calendars.Read` at minimum — these live under
Microsoft Graph, not the `outlook.office.com` audience), kick off an OAuth2
Authorization Code + PKCE flow from a loopback redirect, cache the refresh token, and
call Graph endpoints.

- **Pros**
  - Clean, first-class OAuth2 flow; refresh tokens give us long-lived sessions without
    re-opening a browser on every expiry.
  - Scopes are explicit and auditable; permissions model is the Microsoft-sanctioned
    one.
  - No DOM / web-UI fragility at all.
- **Cons**
  - Most enterprise tenants (banks especially) require **admin consent** to register a
    new application or to grant delegated Graph scopes for a user-owned app. This is
    likely a blocker for the target user.
  - The cleanest APIs are on Graph (`graph.microsoft.com/v1.0/me/messages`), but the
    spec's NG4 explicitly forbids Graph. Using Entra-issued tokens with the
    `outlook.office.com` audience is possible but requires targeting a different
    scope (`https://outlook.office.com/Mail.Read`) and is far less well documented.
  - Adds a heavy runtime dependency (`@azure/msal-node` or hand-rolled PKCE) and a
    config surface (tenant id, client id, redirect URI) that contradicts the spec's
    "no Graph / no app registration" framing.
- **Failure modes**
  - Admin consent denied -> blocked entirely.
  - Refresh-token revocation on password change / MFA reset -> must re-interact.
  - `Mail.Read` scope grant flows sometimes silently downgrade to `Mail.ReadBasic`,
    which omits attachments.
- **Effort: M** (well-trodden path when admin consent is available).

### C. Device Code flow with a well-known public client ID

**What it is:** Reuse a Microsoft-published public client id (e.g. the Azure CLI or the
MS-Teams PowerShell client ids, which are pre-consented in most tenants) and drive the
Device Code OAuth flow to get an access token for `outlook.office.com` or Graph.

- **Pros**
  - No app registration, no admin consent (the public client id is already pre-granted
    in the user's tenant).
  - Nice UX: the user types a short code into a browser on any device.
  - Refresh tokens + long-lived sessions like approach B.
- **Cons**
  - **Terms of service risk.** Microsoft's guidance is clear that these public client
    ids are for their own first-party tools; using them from a third-party CLI is a
    gray-area practice and has been called out in their docs as unsupported. Banks are
    allergic to that.
  - Conditional Access policies frequently block non-enrolled device flows; the user
    may get "device not compliant" errors that the browser-hook approach avoids
    (because the browser session is already compliant).
  - The set of scopes the chosen public-client id will actually mint is fixed and not
    under our control.
- **Failure modes**
  - Conditional Access blocks token issuance.
  - Microsoft rotates / retires the public client id.
  - User's tenant flags the unusual client-id usage and revokes.
- **Effort: M**.

### D. MSAL cache decryption from `localStorage`

**What it is:** Read the `msal.cache.encryption` cookie (AES-GCM key + IV), pull the
encrypted access-token blob from Outlook's IndexedDB / localStorage, decrypt, and use
the plaintext JWT.

- **Pros**
  - No live `fetch` hook required — the token is read at rest.
  - Could be scripted from pure Node with no browser opened once the key + blob are
    known.
  - Potentially survives some UI drift because it does not depend on the login DOM
    sentinel.
- **Cons**
  - **Extremely brittle.** The exact cache layout, key derivation, and encryption
    envelope have changed multiple times in MSAL.js releases (v1 -> v2 -> Browser v3).
  - Microsoft could change the scheme in any OWA release and break us overnight.
  - Spec NG5 explicitly forbids this approach.
- **Failure modes**
  - MSAL.js minor version bump changes the envelope -> silent decryption failure ->
    tool appears to work but returns invalid tokens.
  - Key rotation between browser sessions.
  - Cookie domain / path changes hide the key cookie from us.
- **Effort: L** (reverse-engineering + version-aware parser).

---

## 3. Recommendation

**Primary (this iteration): Approach A — browser-hook Bearer capture + cookie
harvesting.**

The POC already demonstrates end-to-end success, the spec is locked to this path, and
it is the only approach that sidesteps the enterprise admin-consent blocker that kills
approaches B and C for the target user. Approach D is forbidden by NG5 and carries the
highest brittleness anyway.

Risks the implementer must actively handle:

1. **Token expiry within ~1 hour.** Gate every REST call on a 60-second pre-expiry
   check (spec §6.1 state = `expired`). On actual 401, run the re-auth loop exactly
   once before surfacing exit 4.
2. **Header drift.** Replay the minimal set verified in the POC
   (`Authorization: Bearer`, `X-AnchorMailbox: PUID:<puid>@<tenantId>`,
   `Accept: application/json`, cookie jar). If any REST endpoint demands a newer
   header (e.g. `prefer: outlook.body-content-type="text"`), add it per-endpoint
   explicitly; do not blind-copy every header from the browser.
3. **Cookie jar extraction.** Use `context.cookies()` and filter to domains
   `.office.com`, `.outlook.office.com`, `login.microsoftonline.com` at capture time.
   Serialize as `name=value; name2=value2` into the `Cookie:` header, honoring `path`
   and `secure`. Do not drop `httpOnly` cookies — they are required and are visible to
   `context.cookies()` even though they are hidden from `document.cookie`.
4. **First-call race for token capture.** The Node side must subscribe to the
   `exposeBinding` channel _before_ navigating to `outlook.office.com/mail/`, and the
   init-script must be registered via `context.addInitScript` (not `page.addInitScript`)
   so it applies to the very first document including any pre-login redirect chain. The
   capture must return the _first_ Bearer seen; subsequent calls may use the same token
   but there is no value in racing for a "fresher" one.
5. **Browser-closed-mid-capture / MFA timeout.** Wrap the capture in a race against
   `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` and a `page.on('close')` listener; surface exit 4
   without touching the existing session file.

**Secondary (future iteration): Approach B — Entra app registration + PKCE.**

If / when the organization can provision an app registration (or the user has permission
themselves), B is the long-term clean path: refresh tokens eliminate the headed-Chrome
dance, scopes are explicit, and the spec's "no Graph" clause can be relaxed. Track this
in the issues file as a future enhancement.

---

## 4. Implementation strategy (for Approach A)

### 4.1 HTTP client

**Recommendation: native `fetch` (Node ≥ 18), with `AbortController` for the mandatory
per-call timeout.**

Reasoning:

- Zero new dependency; `package.json` stays lean.
- Built-in `AbortController` integrates cleanly with the mandatory
  `OUTLOOK_CLI_HTTP_TIMEOUT_MS` config.
- `undici` would give us pooled keep-alive which is nice for many sequential calls, but
  the CLI issues only 1-2 requests per invocation — not worth the dep.
- `axios` adds a large dep surface and a non-standard response shape. Reject.

### 4.2 CLI framework

**Recommendation: `commander`.**

Reasoning:

- First-class TS types, very small, tree-shake-friendly, ubiquitous in the Node CLI
  ecosystem.
- Clean subcommand syntax maps directly to the 7 commands in the spec.
- `yargs` is more powerful (middleware, positional parsing) but heavier and more
  opinionated; we don't need its extras. `commander` wins on simplicity.

### 4.3 JWT parsing

**Recommendation: manual base64 split.**

Reasoning:

- We only need to read `exp`, `puid`, `tid` from the payload — no signature
  verification (the token is already trusted; we captured it live).
- The parse is 6 lines (`token.split('.')[1]` -> base64url decode -> `JSON.parse`).
- Avoids a dep that we would only call once per capture.
- `jwt-decode` is fine as a fallback if we later need `aud` / `scopes` surfacing, but
  start manual.

### 4.4 Playwright launch model

**Recommendation: `chromium.launchPersistentContext(profileDir, { channel: 'chrome',
headless: false, ... })`.**

Reasoning:

- Spec §7.1 requires a persistent profile directory at
  `$HOME/.outlook-cli/playwright-profile/` (mode 0700) and explicitly treats it as
  sensitive state. `launchPersistentContext` is the direct match.
- `launch + newContext + storageState` works for cookie persistence but **does not
  persist Chrome's session cache**, which is what keeps MSAL silent-SSO working between
  runs. Users would hit MFA prompts on every expiry with `storageState` alone.
- The existing POC script (`test_scripts/outlook_read_recent.ts:102-106`) already uses
  `launchPersistentContext`; we are reusing a known-good pattern.

### 4.5 Concurrency / single-browser guard

**Recommendation: advisory lock file at `$HOME/.outlook-cli/.browser.lock`.**

Design:

- Before launching Playwright, `open(lockPath, 'wx')` (exclusive create). On `EEXIST`,
  read the file: if it contains a PID and that PID is live
  (`process.kill(pid, 0)`), abort with a clear error; otherwise treat as stale and
  overwrite.
- Write `{ pid, startedAt }` JSON into the lock.
- Remove the lock on normal exit and on `SIGINT` / `SIGTERM` / `process.on('exit')`.
- Any command that does **not** open a browser (e.g. `auth-check`, data commands with
  a valid cached session) does not touch the lock.
- Rationale: file-based advisory lock is portable (macOS/Linux/Windows) and dep-free.
  OS-level `flock` is cleaner on POSIX but not portable. Stale detection via `kill -0`
  covers crash recovery (see risk register).

### 4.6 Token capture channel

**Recommendation: `context.addInitScript` that wraps `window.fetch` + a single
`context.exposeBinding('__outlookCliReportAuth', …)` callback.**

Why this over alternatives:

- `page.on('request')` at the Node side can also see `Authorization` headers on modern
  Playwright, but it races with navigation and sometimes omits headers that were added
  by the service-worker. Intercepting in-page is more reliable.
- A DOM event + `page.waitForFunction` is feasible but requires a polling handshake.
  `exposeBinding` gives us a direct in-page -> Node call, fires once per Bearer seen,
  and we resolve our capture promise on the first match.
- The init-script itself: wrap `window.fetch`, read `init.headers` or (if Headers
  instance) `init.headers.get('authorization')`, if the URL starts with
  `https://outlook.office.com/` and the header starts with `Bearer `, call
  `window.__outlookCliReportAuth({ url, auth })`. Keep the wrapper idempotent (guard
  with a `__outlookCliHooked` flag) so HMR / re-navigation doesn't double-wrap.
- We must also handle requests that use a `Request` object (first arg is a `Request`
  instance) — read `req.headers.get('authorization')` off that instance.
- Optionally mirror the same hook over `XMLHttpRequest.prototype.setRequestHeader` for
  defense in depth, but the POC confirms `fetch` is sufficient.

### 4.7 Attachment download

REST shape: `GET /api/v2.0/me/messages/{id}/attachments` returns a JSON collection with
a discriminator field `@odata.type` (or `type` depending on API version) that is one
of:

- `#Microsoft.OutlookServices.FileAttachment` — has `ContentBytes` (base64) plus `Name`,
  `ContentType`, `Size`, `IsInline`. We decode base64 to a `Buffer` and write it.
- `#Microsoft.OutlookServices.ItemAttachment` — an embedded message / event. No
  `ContentBytes`; the payload is a nested `Item` object. **Decision: skip with a
  `reason: "item-attachment"` entry in the `skipped[]` array.** Surfacing embedded
  messages as `.eml` is out of scope for this iteration.
- `#Microsoft.OutlookServices.ReferenceAttachment` — a link to a cloud file
  (OneDrive / SharePoint). No bytes; has a `SourceUrl`. **Decision: skip with a
  `reason: "reference-attachment"` and include the `SourceUrl` in the skip record so
  the user can fetch it manually.**

Flow:

1. `GET /me/messages/{id}/attachments` (list with metadata).
2. For each `FileAttachment`, `GET /me/messages/{id}/attachments/{attId}` to fetch the
   full payload including `ContentBytes`. (The list endpoint sometimes omits the bytes
   for large attachments; always fetch the detail endpoint.)
3. `--include-inline=false` (default): skip any with `IsInline: true`; record in
   `skipped[]` with `reason: "inline"`.
4. `--overwrite=false` (default): if the target path exists, exit 6 with a clear
   message listing the offending file.
5. Write via §4.8 below.

### 4.8 Atomic file writes with mode 0600

**Recommendation: `fs.openSync(tmpPath, 'wx', 0o600)` -> `fs.writeSync(fd, buf)` ->
`fs.closeSync(fd)` -> `fs.renameSync(tmpPath, finalPath)`.**

Details:

- `wx` flag fails if `tmpPath` already exists, giving us a natural guard against
  stale temp files from crashed runs (retry with a fresh random suffix).
- Mode `0o600` is applied at `open` time, so the file is never readable by other users
  even for an instant (writing then `chmod`ing leaves a race window).
- `rename` is atomic on the same filesystem. Keep `tmpPath` in the same directory as
  `finalPath` to guarantee that.
- For the session dir itself: `fs.mkdirSync(dir, { recursive: true, mode: 0o700 })` —
  then `fs.chmodSync(dir, 0o700)` defensively in case `recursive` created intermediate
  dirs with umask-dependent modes.

### 4.9 Error taxonomy and exit codes

Map from upstream signal -> our exit code / error class:

| Signal                                     | Exit    | Class                | Notes                                                                                   |
| ------------------------------------------ | ------- | -------------------- | --------------------------------------------------------------------------------------- |
| HTTP 401 on first attempt                  | (retry) | —                    | Triggers re-auth exactly once.                                                          |
| HTTP 401 on second attempt                 | 4       | `AuthError`          | Do not re-open browser again.                                                           |
| HTTP 403                                   | 5       | `UpstreamError`      | Includes "conditional access" denials. No retry.                                        |
| HTTP 404 (bad id)                          | 5       | `UpstreamError`      | Surface `requestedId` in error payload.                                                 |
| HTTP 429                                   | 5       | `UpstreamError`      | Include `Retry-After` if present; no auto-retry in this iteration.                      |
| HTTP 5xx                                   | 5       | `UpstreamError`      | No auto-retry; user can re-run.                                                         |
| Network / DNS / TLS                        | 5       | `UpstreamError`      | Wrap the underlying Error; do not leak the token in any thrown error's `request` field. |
| AbortError (timeout)                       | 5       | `UpstreamError`      | Message: "HTTP timeout after Nms".                                                      |
| Session file read error                    | 6       | `IoError`            | `EACCES`, `ENOENT` on parent, etc.                                                      |
| Session file write error                   | 6       | `IoError`            | Same.                                                                                   |
| Attachment target exists w/o `--overwrite` | 6       | `IoError`            | Names the offending file.                                                               |
| Missing mandatory config                   | 3       | `ConfigurationError` | Names the setting + full precedence chain checked.                                      |
| Invalid argv                               | 2       | (commander default)  | `commander` handles this natively.                                                      |
| User closes browser / login timeout        | 4       | `AuthError`          | Message: "login not completed".                                                         |

Every thrown error includes `{ code, exitCode, cause? }` and is caught by a top-level
`main()` that does the JSON / stderr formatting and `process.exit(err.exitCode)`.

---

## 5. Risk register

| Risk                                                                                             | Likelihood | Impact | Mitigation                                                                                                                                                                                                                                                                         |
| ------------------------------------------------------------------------------------------------ | ---------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Bearer token expires between `auth-check` and the real REST call                                 | High       | Medium | Gate on `now + 60 s >= expiresAt` per §6.1; on 401, auto re-auth + retry once (§6.4).                                                                                                                                                                                              |
| MFA re-prompted on silent re-login (e.g. expired conditional-access session)                     | Medium     | Medium | Headed Chrome stays open for up to `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`; user completes MFA normally. Never silence MFA UI.                                                                                                                                                              |
| Outlook UI changes break the "inbox reached" DOM sentinel                                        | Medium     | High   | Detect by URL regex first (`^https://outlook.office.com/mail/`), DOM sentinel second. Keep the DOM probe loose (`div[role="main"], div[role="option"]`). Also treat "first captured Bearer" as a sufficient signal — if we already have a token, we don't _need_ DOM confirmation. |
| Captured Bearer is missing a required scope (e.g. no `Calendars.Read`)                           | Low-Medium | High   | The OWA Bearer ships with scopes for the full Outlook web surface; calendar scopes are included. Validate with an `auth-check` that hits both `/me/messages?$top=1` and `/me/calendarview?...&$top=1` and surfaces the first failure.                                              |
| Cookie domain mismatch (harvested cookies from `.office.com` don't satisfy `outlook.office.com`) | Low        | High   | Filter `context.cookies()` to `.office.com`, `.outlook.office.com`, `.login.microsoftonline.com` — broad enough to cover all paths; cookie library honors domain suffix matching per RFC 6265 when serialized. Verify in test.                                                     |
| Lock file stale after crash / SIGKILL                                                            | Medium     | Low    | Store PID in lock; on conflict, `process.kill(pid, 0)` — if it throws `ESRCH`, treat as stale and overwrite. Also expire locks older than `max(login_timeout, 30 min)`.                                                                                                            |
| Browser window closed mid-capture (user X's out)                                                 | Medium     | Low    | Race capture promise against `page.on('close')` and `context.on('close')`; reject with `AuthError("login not completed")` -> exit 4. Do not touch existing session file.                                                                                                           |
| First API request is not a `fetch` call (e.g. done by a service worker in a way we don't see)    | Low        | Medium | Secondary `XMLHttpRequest` hook in the init-script as a defense-in-depth; additionally, `page.on('request')` listener in Node as a tertiary fallback. Surface a clear error if no Bearer is captured within the login timeout.                                                     |
| Secrets leak via error messages / log file                                                       | Low        | High   | `ConfigurationError`, `UpstreamError`, `AuthError` all sanitize: the Bearer token and cookie values are never included in `.message`, `.stack`, or any log record. AC-NO-SECRET-LEAK covers this.                                                                                  |
| Persistent profile corrupted on disk (partial write / disk full)                                 | Very Low   | Medium | Spec allows `--force` to rebuild; document that a user can delete `$HOME/.outlook-cli/playwright-profile/` to reset. Do not attempt self-repair.                                                                                                                                   |

---

## 6. Technical Research Guidance

```
Research needed: Yes

Topic: Playwright addInitScript + exposeBinding for reliable first-call Bearer capture
Why: The POC worked, but we need the minimum-drift pattern (context-level init script, fetch + Request-object handling, XMLHttpRequest fallback, exact shape of exposeBinding invocation) documented before we commit to `src/auth/fetchHook.ts`.
Focus: context.addInitScript vs page.addInitScript scoping, wrapping window.fetch when init.headers is a Headers instance vs plain object vs Request object, exposeBinding call semantics + once-only resolution, timing of init-script registration relative to goto, Playwright 1.59 specifics
Depth: medium

Topic: outlook.office.com/api/v2.0 attachment download semantics (FileAttachment vs ItemAttachment vs ReferenceAttachment)
Why: Spec §5.5 requires correct behavior on every attachment subtype; we proposed skipping ItemAttachment / ReferenceAttachment but need to confirm the exact @odata.type discriminator, whether the detail endpoint differs per type, and whether `$value` is supported on v2 (it is on Graph; ambiguous on v2).
Focus: @odata.type discriminator values in v2 responses, presence of ContentBytes on list vs detail, $value support, ReferenceAttachment SourceUrl field, size limits that force a different endpoint
Depth: medium
```

---

## Summary

- **Primary approach: A (browser-hook).** POC-validated, no admin consent needed, fits
  the enterprise / bank context. Risks are well understood and bounded.
- **Secondary (future): B (Entra app + PKCE)** once the org allows app registration;
  promises refresh-token-driven silent operation.
- **Implementation picks:** native `fetch` + `AbortController`, `commander`, manual JWT
  base64 split, `launchPersistentContext` with `channel: 'chrome'`, advisory PID lock
  under `$HOME/.outlook-cli/`, `addInitScript` + `exposeBinding` for token capture,
  `open(0o600) + rename` for atomic writes.
- **Open research:** two focused topics — Playwright capture pattern and Outlook v2
  attachment subtypes — flagged for Phase 3b before implementation begins.

Absolute output path: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/design/investigation-outlook-cli.md`
