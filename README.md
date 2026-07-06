# outlook-cli

[![CI](https://github.com/weirdapps/outlook-access/actions/workflows/ci.yml/badge.svg)](https://github.com/weirdapps/outlook-access/actions/workflows/ci.yml)
[![CodeQL](https://github.com/weirdapps/outlook-access/actions/workflows/codeql.yml/badge.svg)](https://github.com/weirdapps/outlook-access/actions/workflows/codeql.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Node.js 22](https://img.shields.io/badge/node-%3E%3D20-brightgreen.svg)](https://nodejs.org/)
[![Version](https://img.shields.io/badge/version-1.5.0-blue.svg)](package.json)

Drive your personal Microsoft 365 / Outlook mailbox from the command line without
registering an app, without asking your tenant admin, and without provisioning
Graph OAuth clients.

`outlook-cli` opens a real Chrome window (via Playwright) once, watches the
outbound requests that Outlook Web already makes, and grabs the `Authorization:
Bearer ...` token plus session cookies. That capture is written atomically to
`~/.outlook-cli/session.json` (file mode 0600, parent directory mode 0700) and
replayed against the same `https://outlook.office.com/api/v2.0/` REST surface
your browser talks to. After the first login, every subsequent command is
scriptable and non-interactive; when the token expires the tool silently
re-opens the browser (headless if the persistent profile is still warm) unless
you pass `--no-auto-reauth`.

This is the source repo for the `outlook-cli` binary. It is also consumed as a
git-pinned npm dependency by the `outlook-bridge` MCP server in the
[plessas-marketplace `mail` plugin](https://github.com/weirdapps/plessas-marketplace/tree/master/plugins/mail),
which surfaces the same capabilities to Claude Code. Sister project:
[`teams-access`](https://github.com/weirdapps/teams-access), which applies the
identical web-session capture pattern to Microsoft Teams.

## Why this exists

The usual ways to script against a personal M365 mailbox are all heavy for a
single-user use case:

- **Microsoft Graph** wants an app registration, admin consent, an OAuth client,
  `Mail.*` / `Calendars.*` scopes, and a redirect-URI story.
- **EWS / MAPI** is deprecated, Windows-flavoured, and largely out of tenant.
- **IMAP / SMTP** is usually disabled by policy in modern M365 tenants.

If you can sign in to `outlook.office.com` in a browser, you can already reach
the mailbox. `outlook-cli` reuses that fact: sign in once, capture the token,
replay it. Nothing bypasses conditional access, MFA, or tenant policy; you go
through them exactly as the browser would.

## Commands

Nineteen subcommands are wired up in `src/cli.ts`. Every command emits JSON on
stdout by default and accepts `--table` for a compact human view. Errors are
always emitted as JSON on stderr with a `code` field and a numeric exit code
(see [Exit codes](#exit-codes)).

| Command                          | Purpose                                                                                                                                                                                                                                                                                                                            | Backend             |
| -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------- |
| `login`                          | One-shot headed Chrome window; captures the first outbound Bearer + cookies and writes `session.json`. `--sharepoint-host <host>` also captures a SharePoint session. `--force` ignores any cached session.                                                                                                                        | Playwright          |
| `auth-check`                     | Non-interactive verification that the cached session is still accepted.                                                                                                                                                                                                                                                            | outlook.office.com  |
| `auth-renew`                     | Silent (headless) bearer refresh using the persisted browser profile. `--sharepoint-host` also refreshes a SharePoint session.                                                                                                                                                                                                     | Playwright headless |
| `list-mail`                      | List messages from a folder. Supports `--folder` / `--folder-id` / `--folder-parent`, `--since` / `--until` or the keyword-aware `--from` / `--to`, `--all` pagination with `--max` safety cap, `--select`, and `--just-count` (server-side `$count=true`).                                                                        | REST v2.0 `$filter` |
| `get-mail <id>`                  | Retrieve one message. `--body html\|text\|none`.                                                                                                                                                                                                                                                                                   | REST v2.0           |
| `get-thread <id>`                | Retrieve every message in a conversation, across folders. Accepts `conv:<conversationId>` to skip the resolve hop. `--order asc\|desc`, `--body`.                                                                                                                                                                                  | REST v2.0           |
| `download-attachments <id>`      | Save non-inline attachments to `--out <dir>`. `--include-inline`, `--overwrite`.                                                                                                                                                                                                                                                   | REST v2.0           |
| `download-sharepoint-link <url>` | Fetch a `ReferenceAttachment.SourceUrl` (SharePoint / OneDrive-for-Business) using the captured SharePoint session.                                                                                                                                                                                                                | SharePoint REST     |
| `list-calendar`                  | List events in a window. `--from` / `--to` accept ISO-8601 or keywords (`now`, `now + 7d`).                                                                                                                                                                                                                                        | REST v2.0           |
| `get-event <id>`                 | Retrieve one calendar event. `--body`.                                                                                                                                                                                                                                                                                             | REST v2.0           |
| `list-folders`                   | List folders under a parent (well-known alias, path, or `id:<raw>`). `--recursive`, `--include-hidden`, `--first-match`.                                                                                                                                                                                                           | REST v2.0           |
| `find-folder <spec>`             | Resolve a folder query (alias, path, or `id:<raw>`) to a single `ResolvedFolder`. `--anchor`, `--first-match`.                                                                                                                                                                                                                     | REST v2.0           |
| `create-folder <path>`           | Create (or idempotently reuse) a mail folder. `--parent`, `--create-parents`, `--idempotent`.                                                                                                                                                                                                                                      | REST v2.0           |
| `move-mail <ids...>`             | Move one or more messages to `--to <spec>`. `--continue-on-error` collects failures instead of aborting; per-message failures still set exit 5.                                                                                                                                                                                    | REST v2.0           |
| `send-mail`                      | Compose and send. Default: creates a draft and activates Outlook desktop (macOS only). `--send-now` dispatches immediately. `--to` / `--cc` / `--bcc`, `--subject`, `--html` / `--text`, `--attach` (repeatable, combined cap 30 MB), `--signature`, `--no-signature`, `--no-cc-self`, `--no-save-sent`, `--no-open`, `--dry-run`. | REST v2.0           |
| `capture-signature`              | Extract a signature from a SentItems message and save to `~/.outlook-cli/signature.html`. `--from-message <id>`, `--out <file>`.                                                                                                                                                                                                   | REST v2.0           |
| `reply <id>`                     | Reply to a message. Auto-quotes original, appends signature. Same draft-first / `--send-now` model as `send-mail`.                                                                                                                                                                                                                 | REST v2.0           |
| `reply-all <id>`                 | Reply-all. Recipients are pre-populated by M365.                                                                                                                                                                                                                                                                                   | REST v2.0           |
| `forward <id>`                   | Forward a message. `--to` required. Auto-quotes original.                                                                                                                                                                                                                                                                          | REST v2.0           |

`outlook-cli <command> --help` prints the full flag set for each subcommand.

## How it works

```mermaid
flowchart LR
    A["User"] -->|"outlook-cli login"| B["Playwright<br/>headed Chrome"]
    B -->|"sign in normally<br/>(MFA / conditional access)"| C["outlook.office.com"]
    B -->|"snoop first<br/>Authorization: Bearer"| D["captureOutlookSession"]
    D -->|"atomic write<br/>mode 0600"| E[("~/.outlook-cli/<br/>session.json")]
    D -.optional.-> F[("~/.outlook-cli/<br/>sharepoint-session.json")]
    E --> G["OutlookClient<br/>(src/http)"]
    G -->|"fetch + Bearer<br/>+ cookies"| H["outlook.office.com<br/>/api/v2.0/*"]
    G -->|"on 401,<br/>if !noAutoReauth"| B
    I["list-mail, send-mail,<br/>reply, forward, folders,<br/>calendar, ..."] --> G
```

## Requirements

- **Node.js 20 LTS or newer.** CI runs Node 22. Older Node lacks global `fetch`
  and other APIs the tool depends on.
- **npm 10+** (bundled with modern Node). `package-lock.json` is committed; no
  yarn / pnpm support is assumed.
- **A real Google Chrome or Microsoft Edge install on your machine.** Playwright
  launches your installed browser via the channel mechanism
  (`chromium.launchPersistentContext({ channel })`); it does not download its
  own Chromium build. Accepted channel values: `chrome` (default),
  `chrome-beta`, `chrome-dev`, `msedge`, `msedge-beta`.
- **A Microsoft 365 / Office 365 mailbox** you can sign in to at
  `outlook.office.com`. Consumer `outlook.live.com` / `hotmail.com` mailboxes
  use a different API surface and are not supported.
- Outbound HTTPS to `outlook.office.com`, `login.microsoftonline.com`, and
  whatever conditional-access endpoints your tenant routes through.
- Write access to `$HOME` (the session file lives at
  `$HOME/.outlook-cli/session.json`). POSIX file-mode enforcement is strict on
  macOS and Linux. On Windows the file is still written atomically but ACL
  hardening is your responsibility.

## Install and build

```bash
git clone https://github.com/weirdapps/outlook-access.git
cd outlook-access
npm install
npm run build          # emits dist/cli.js (chmod +x in postbuild)
```

Link the CLI globally so you can invoke it as `outlook-cli` from anywhere:

```bash
npm link               # installs a symlink at $(npm prefix -g)/bin/outlook-cli
```

Or run the TypeScript sources directly without building:

```bash
npx ts-node src/cli.ts <subcommand> [options]
# or
npm run cli -- <subcommand> [options]
```

## First use

```bash
outlook-cli login
```

A Chrome window opens at `https://outlook.office.com/`. Sign in normally,
completing whatever MFA or conditional-access step your tenant requires. The
tool watches outbound requests, captures the first `Authorization: Bearer`
header it sees, closes the window, and writes `~/.outlook-cli/session.json`.

Verify:

```bash
outlook-cli auth-check
# {
#   "status": "ok",
#   "tokenExpiresAt": "2026-04-22T15:03:25.000Z",
#   "account": { "upn": "you@yourtenant.com" }
# }
```

After that, every subcommand replays the cached session. When the token
expires, the default behaviour is to re-open the browser silently and refresh
it. Pass `--no-auto-reauth` if you want expired-session failures to be hard
errors (exit 4) instead.

To also capture a SharePoint session for `download-sharepoint-link`:

```bash
outlook-cli login --sharepoint-host <tenant>.sharepoint.com
# writes ~/.outlook-cli/sharepoint-session.json
```

## Usage

### Mail

```bash
# Most recent 5 inbox messages, as a table
outlook-cli list-mail --top 5 --table

# A specific message, body as text
outlook-cli get-mail AAMkAGI... --body text > message.json

# Save all non-inline attachments to ./att
outlook-cli download-attachments AAMkAGI... --out ./att

# Incremental sync: last 24h, paginated, capped at 5000
outlook-cli list-mail \
  --folder Inbox \
  --since "$(date -u -v-24H +%Y-%m-%dT%H:%M:%SZ)" \
  --all --max 5000 --json

# Explicit date window
outlook-cli list-mail \
  --since 2026-04-01T00:00:00Z \
  --until 2026-04-08T00:00:00Z \
  --all --json

# Server-side count only (no message rows)
outlook-cli list-mail --folder Inbox --just-count

# Full thread across folders (or pass "conv:<id>" to skip the resolve hop)
outlook-cli get-thread AAMkAGI... --order asc
```

`--since` / `--until` add a server-side `$filter` on `ReceivedDateTime`. The
newer `--from` / `--to` accept the same ISO-8601 plus keywords (`now`,
`now+7d`, `now-24h`). `--all` walks `@odata.nextLink` until exhausted;
`--max <N>` is the safety cap (default 10000, max 100000). When the cap is
hit and more results remain, a `max_results_reached` warning is emitted on
stderr and the partial result is returned.

### Send, reply, forward

```bash
# Compose new email; default creates a draft and activates Outlook desktop (macOS)
outlook-cli send-mail \
  --to "alice@example.com" "bob@example.com" \
  --cc "carol@example.com" \
  --subject "Q2 review" \
  --html body.html

# Skip the draft, send immediately
outlook-cli send-mail \
  --to "alice@example.com" \
  --subject "quick update" \
  --html body.html \
  --send-now

# Attach files (combined cap 30 MB)
outlook-cli send-mail \
  --to "alice@example.com" \
  --subject "report attached" \
  --html body.html \
  --attach report.pdf --attach slides.pptx

# Reply to a message (auto-quotes original, appends signature)
outlook-cli reply AAMkAGI... --html reply.html

# Reply-all
outlook-cli reply-all AAMkAGI... --html reply.html --send-now

# Forward (--to required)
outlook-cli forward AAMkAGI... \
  --to "dave@example.com" \
  --html note.html

# Extract your signature from the latest SentItems message
outlook-cli capture-signature
outlook-cli capture-signature --from-message AAMkAGI...
```

All send / reply / forward commands default to draft-first: the message is
created as a draft and Outlook desktop is activated so you can review before
sending. Pass `--send-now` to dispatch immediately. Automatic CC-self is on
by default (suppress with `--no-cc-self`). Signature from
`~/.outlook-cli/signature.html` is appended automatically (suppress with
`--no-signature`). Outlook-desktop activation is macOS-only; on Linux or
Windows the draft is still created, `--no-open` becomes a no-op, and a
`skipping (platform=..., only darwin is supported)` note is written to stderr.

### Calendar

```bash
# Next 14 days
outlook-cli list-calendar --from now --to "now + 14d" --table

# One event
outlook-cli get-event AAMkAGI...
```

### Folders

```bash
# Top-level folders
outlook-cli list-folders --table

# Full sub-tree (bounded)
outlook-cli list-folders --recursive --table

# Resolve a folder by path
outlook-cli find-folder "Inbox/Projects/Alpha"

# Create nested folder idempotently
outlook-cli create-folder "Inbox/Projects/Alpha" --create-parents --idempotent

# Move messages (by alias, path, or "id:<raw>")
outlook-cli move-mail AAMk... AAMk... --to "Inbox/Archive-2026"
outlook-cli move-mail AAMk... --to Archive
outlook-cli move-mail AAMk... --to "id:AAMkAGI..." --continue-on-error
```

### SharePoint reference attachments

Some Outlook messages carry `ReferenceAttachment` entries: SharePoint or
OneDrive-for-Business shared links rather than inline binaries. Their content
lives on `<tenant>.sharepoint.com`, which uses a different Bearer token and
cookies from `outlook.office.com`. Capture that second session at login time:

```bash
outlook-cli login --sharepoint-host <tenant>.sharepoint.com
# writes ~/.outlook-cli/sharepoint-session.json (mode 0600)

outlook-cli download-sharepoint-link \
  "https://<tenant>.sharepoint.com/sites/foo/Documents/report.pdf" \
  --out ./att
```

If the SharePoint session file is missing or expired, the command exits with
code 4 and prints the exact `outlook-cli login` invocation to recover.

### List mail from an arbitrary folder

```bash
outlook-cli list-mail --folder "Inbox/Projects/Alpha" --top 10 --table
outlook-cli list-mail --folder-id AAMkAGI... --top 20
outlook-cli list-mail --folder-parent Inbox --folder "Projects/Alpha"
```

## Output modes

Every subcommand supports two formats, mutually exclusive:

- `--json` (default), stable, stdout, pipe into `jq` or scripts.
- `--table`, human-readable, compact columns. IDs are never truncated so they
  can be copy-pasted back into other subcommands.

Errors are always emitted as JSON on stderr with `code`, optional `message`,
and setting-specific fields (e.g. `missingSetting`, `path`, `failed[]`).

## Configuration

`outlook-cli` needs no configuration file for a basic install. The runtime
plumbing has three tunable settings, each with a default:

| Setting                        | CLI flag                  | Env var                        | Default          |
| ------------------------------ | ------------------------- | ------------------------------ | ---------------- |
| Per-REST-call HTTP timeout     | `--timeout <ms>`          | `OUTLOOK_CLI_HTTP_TIMEOUT_MS`  | `30000` (30 s)   |
| Max wait for interactive login | `--login-timeout <ms>`    | `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` | `300000` (5 min) |
| Playwright Chrome channel      | `--chrome-channel <name>` | `OUTLOOK_CLI_CHROME_CHANNEL`   | `chrome`         |

Additional overrides:

| Setting                | CLI flag                | Env var                    | Default                             |
| ---------------------- | ----------------------- | -------------------------- | ----------------------------------- |
| Session file path      | `--session-file <path>` | `OUTLOOK_CLI_SESSION_FILE` | `~/.outlook-cli/session.json`       |
| Playwright profile dir | `--profile-dir <path>`  | `OUTLOOK_CLI_PROFILE_DIR`  | `~/.outlook-cli/playwright-profile` |
| IANA timezone          | `--tz <iana>`           | `OUTLOOK_CLI_TZ`           | `process.env.TZ` or system tz       |
| Calendar window start  | `--from <iso\|keyword>` | `OUTLOOK_CLI_CAL_FROM`     | `now`                               |
| Calendar window end    | `--to <iso\|keyword>`   | `OUTLOOK_CLI_CAL_TO`       | `now + 7d`                          |

Precedence: CLI flag beats env var beats default. A malformed flag or env
value still throws `ConfigurationError` (exit 3); the default only covers the
unset case. For persistent overrides, `source ./outlook-cli.env` in your shell
or append it to `~/.zshrc` / `~/.bashrc`.

### Runtime data under `~/.outlook-cli/`

| Path                                  | Written by                                                | Purpose                                                                                             |
| ------------------------------------- | --------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `session.json` (mode 0600)            | `login`, `auth-renew`, silent re-auth                     | Bearer token, cookies, account UPN. Read by every REST call.                                        |
| `sharepoint-session.json` (mode 0600) | `login --sharepoint-host`, `auth-renew --sharepoint-host` | SharePoint Bearer + cookies for `download-sharepoint-link`.                                         |
| `playwright-profile/` (mode 0700)     | Playwright persistent context                             | Persists browser state so silent re-auth works headless.                                            |
| `signature.html`                      | `capture-signature`                                       | Optional. Auto-appended by `send-mail` / `reply` / `reply-all` / `forward` unless `--no-signature`. |
| `signature-assets/`                   | `capture-signature` (as needed)                           | Inline images extracted from the signature.                                                         |

Nothing in `~/.outlook-cli/` is ever printed or logged. Body-snippet redaction
(`src/util/redact.ts`) runs on every error path.

### Exit codes

| Code | Meaning                                                                                                   |
| ---- | --------------------------------------------------------------------------------------------------------- |
| `0`  | Success                                                                                                   |
| `1`  | Unexpected error                                                                                          |
| `2`  | Invalid usage (bad argv, commander error)                                                                 |
| `3`  | Configuration error (malformed flag or env var)                                                           |
| `4`  | Auth failure (expired / rejected session, user cancelled login, `--no-auto-reauth` with no cache)         |
| `5`  | Upstream API error (non-401 HTTP error, timeout, network failure, pagination limit, partial move failure) |
| `6`  | IO error (folder collision without `--idempotent`, file collision without `--overwrite`)                  |

## Architecture

```text
src/
  cli.ts                    Commander wiring, global options, error mapping
  auth/
    browser-capture.ts      Playwright login + first-Bearer capture
    sharepoint-capture.ts   SharePoint token capture (same browser context)
    jwt.ts                  Base64-URL decode + expiry parsing
    lock.ts                 Session-file locking
  session/
    schema.ts               SessionFile type + validation
    sharepoint-schema.ts    Sharepoint session file type + IO
    store.ts                Atomic read / write with fs-atomic
  http/
    outlook-client.ts       fetch() wrapper: Bearer + cookies + 401 re-auth
    sharepoint-client.ts    SharePoint fetch() wrapper
    filter-builder.ts       Server-side $filter helpers
    errors.ts               UpstreamError, CollisionError, ...
    types.ts                MessageSummary, EventSummary, FolderSummary, ...
  folders/
    resolver.ts             Well-known alias / path / id:<raw> resolution
    types.ts                ResolvedFolder, CreateFolderResult, MoveMailResult
  commands/                 One file per subcommand
  config/
    config.ts               loadConfig with flag > env > default precedence
    errors.ts               ConfigurationError, AuthError, IoError, ...
  output/
    formatter.ts            JSON and table rendering (ColumnSpec)
  util/
    dates.ts                ISO-8601 + keyword parsing (now, now+7d)
    filename.ts             Safe attachment filenames
    fs-atomic.ts            Atomic tmpfile + rename + chmod
    open-outlook.ts         macOS `open -a Microsoft Outlook` shim
    redact.ts               Body-snippet redaction on every error path
    signature-assets.ts     Inline image extraction for signatures
test_scripts/               vitest suites (see below)
docs/
  design/                   project-design.md, plan-NNN-*.md, refined-request-*.md, configuration-guide.md
  reference/                Codebase scans
  research/                 Outlook REST v2.0 quirks
  superpowers/              Repo-automation workflow notes
scripts/
  pii-gauntlet.sh           Grep gauntlet for accidental PII in fixtures / docs
```

### Runtime dependencies

Runtime footprint is deliberately small: one direct dependency for CLI
parsing, one optional dependency for the login browser.

| Package                                           | Version              | Role                                                                                                                                      |
| ------------------------------------------------- | -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| [`commander`](https://github.com/tj/commander.js) | `^14.0.3`            | CLI parser: subcommands, options, help output                                                                                             |
| [`playwright`](https://playwright.dev/)           | `^1.59.1` (optional) | Drives the headed Chrome window during `login` and captures the outbound Bearer. Lazy-loaded so read-only commands skip the browser init. |

Everything else (HTTP, JSON, file IO, crypto, timezone math, JWT decode) uses
Node's built-in `node:*` modules.

## Development

```bash
npm install                # installs deps and playwright as an optional dep
npm run build              # tsc; emits dist/cli.js and chmod +x
npm run lint               # eslint .
npm test                   # vitest run (394 tests across 34 files, at v1.5.0)
npm run test:watch         # incremental vitest
npm run test:coverage      # vitest run --coverage (v8)
npm run format             # prettier --write .
```

CI (`.github/workflows/ci.yml`) runs `npm ci --ignore-scripts`, then lint,
build, and test on Node 22. CodeQL scans on push, PR, and weekly cron.
SonarCloud runs on push and PR. Dependabot manages npm and GitHub Actions
updates with an auto-merge workflow for patch and minor; a monthly
`deps-refresh` workflow (via `weirdapps/shared-workflows`) opens a
consolidated PR when lock-only updates are available.

### PII gauntlet

`scripts/pii-gauntlet.sh` greps fixtures, docs, and source for accidentally
committed personal data. Run it before opening a PR that touches test
fixtures or documentation.

## Consumed by

- **`outlook-bridge` MCP server** in the [plessas-marketplace `mail` plugin](https://github.com/weirdapps/plessas-marketplace/tree/master/plugins/mail).
  It installs `outlook-tool` from this repo as a git-pinned dependency and
  exposes the same command surface to Claude Code as MCP tools
  (`outlook_list_mail`, `outlook_send_mail`, `outlook_reply`, ...).
- **Sister project**: [`teams-access`](https://github.com/weirdapps/teams-access)
  applies the identical web-session capture pattern to Microsoft Teams.

## Security posture

The session file contains a live Bearer token and cookies. It is written
atomically under a `0700` directory with mode `0600`, is never printed or
logged (body-snippet redaction runs on every error path), and is
`.gitignore`d alongside the Playwright profile directory. See
[`SECURITY.md`](SECURITY.md) for disclosure policy.

## Origin

Forked from [BikS2013/outlook-tool](https://github.com/BikS2013/outlook-tool)
by Giorgos Marinos, whose core insight (capturing an Outlook Web bearer via
headed Playwright and reusing it against `outlook.office.com/api/v2.0`) made
this approach viable. The codebase has since been substantially rewritten and
extended: folder management, send / reply / forward with signature and
inline-image support, silent token renewal, SharePoint reference attachments,
atomic session storage with file locking, redaction on every error path, and
the current 394-test vitest suite.

## License

MIT. See [LICENSE](LICENSE) for full text and the dual copyright covering the
original upstream and this fork's substantial rewrite.
