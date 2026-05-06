# outlook-cli

A TypeScript command-line tool for reading, sending, and organizing Outlook
mail and calendar, by reusing an interactively captured Outlook-web session
— no app registration, no tenant admin, no API keys.

---

## Rationale

The usual ways to script against an Outlook/Exchange mailbox are all heavy:

- **Microsoft Graph** requires an app registration, admin consent, OAuth client
  credentials, a tenant willing to grant `Mail.*` / `Calendars.*` scopes, and a
  working redirect-URI plumbing. Fine for a service; overkill for one person who
  just wants to read their own inbox from a script.
- **EWS / MAPI** is deprecated, on-prem-flavored, and Windows-centric.
- **IMAP/SMTP** is usually disabled in modern tenants.

For a single user who is already allowed to sign in to `outlook.office.com` in
a browser, there is a much shorter path: log in **once** in a headed Chrome
window, grab the Bearer token and cookies that the web UI itself uses, cache
them securely, and drive the same **`https://outlook.office.com/api/v2.0/...`**
REST surface that Outlook-web talks to.

That is what this tool does.

### What it gives you

- `login` — a one-shot headed Playwright Chrome window that you use to sign in
  normally (including MFA / conditional access). The tool snoops the first
  outbound request bearing `Authorization: Bearer`, extracts the token +
  cookies, and writes them atomically to `$HOME/.outlook-cli/session.json`
  (mode `0600`, parent dir `0700`).
- `auth-check` — non-interactive verification that the cached session is still
  accepted.
- `list-mail`, `get-mail`, `get-thread`, `download-attachments` — inbox +
  message access, including full conversation threading and saving attachments
  to a directory.
- `send-mail` — compose and send new emails with HTML body, attachments, and
  automatic signature. Default: creates a draft and activates Outlook desktop;
  `--send-now` dispatches immediately.
- `reply`, `reply-all`, `forward` — respond to or forward messages with
  auto-quoted original content and signature. Same draft-first default.
- `capture-signature` — extract your email signature from a sent message and
  save to `~/.outlook-cli/signature.html` for automatic appending.
- `list-calendar`, `get-event` — calendar window listing and single-event retrieval.
- `list-folders`, `find-folder`, `create-folder`, `move-mail` — full folder
  management: list, resolve by name/path/id, create (idempotently), and move
  messages across folders.
- Every subcommand re-uses the cached session. When the token expires the tool
  auto-reopens the Playwright window for a silent re-auth (unless
  `--no-auto-reauth` is passed, which makes expired-session failures hard).

### What it deliberately does **not** do

- It does not delete messages or modify calendar events.
- It does not persist anything upstream beyond sent messages. The session file
  is local-only.
- It does not bypass conditional access, MFA, or any tenant policy — you log in
  exactly the way the browser would.

### Security posture in one line

The session file contains a live Bearer token + cookies. It is written atomically
under a 0700 directory with mode 0600, is never printed or logged (body-snippet
redaction runs on every error path), and is `.gitignore`d alongside the
Playwright profile dir.

---

## Prerequisites

### Runtime environment

- **Node.js 20 LTS or newer** (tested on 22). Older Node versions lack the
  global `fetch` and other APIs the tool relies on.
- **npm 10+** (bundled with Node 20+). The repo ships a `package-lock.json`; no
  yarn/pnpm support is assumed.
- **git** — only needed to clone the repo.

### Browser (critical)

- A **real Google Chrome or Microsoft Edge installation on your machine**.
  Playwright launches your _installed_ browser via the `channel` mechanism
  (`chromium.launchPersistentContext({ channel: ... })`), it does **not**
  download its own Chromium build. You therefore do **not** need to run
  `npx playwright install`.
- Accepted channel values: `chrome` (default), `chrome-beta`, `chrome-dev`,
  `msedge`, `msedge-beta`. Whichever you pick must actually be installed and
  locatable by Playwright.
- macOS, Linux, and Windows are all supported so long as the chosen channel
  exists on the system.

### Network

- Outbound HTTPS to `https://outlook.office.com/*` and the Microsoft sign-in
  chain (`login.microsoftonline.com`, conditional-access endpoints your tenant
  routes through, etc.).
- No inbound ports are opened. No proxy is configured; use your system proxy.

### Account

- A **Microsoft 365 / Office 365 mailbox** you can sign into at
  `outlook.office.com` (work or school, or a personal MSA that resolves to
  that endpoint). Legacy `outlook.live.com` / `hotmail.com` consumer mailboxes
  use a different API surface and are **not** supported by this tool.
- Your tenant's conditional-access / MFA policies apply exactly as they would
  in the browser — you complete them in the Playwright window during `login`.

### Platform / permissions

- Write access to `$HOME` (the session file lives at
  `$HOME/.outlook-cli/session.json`, parent dir `0700`, file `0600`).
- On **macOS / Linux**, the POSIX file-mode enforcement is strict.
- On **Windows**, `fs.chmod` is largely a no-op — the session file is still
  written atomically, but filesystem ACL hardening is your responsibility.
  The rest of the tool (Playwright, HTTP, commander wiring) works unchanged.

---

## Libraries & tools used

### Runtime (production)

| Package                                           | Version   | Why                                                   |
| ------------------------------------------------- | --------- | ----------------------------------------------------- |
| [`commander`](https://github.com/tj/commander.js) | `^14.0.3` | CLI parser — subcommands, option mixing, help output. |

That is the entire runtime footprint. Everything else (HTTP, JSON parsing,
file IO, crypto, timezone math, token base64-URL decoding) uses Node's
built-in `node:*` modules.

### Development / build / test

| Package                                                    | Version   | Why                                                                                              |
| ---------------------------------------------------------- | --------- | ------------------------------------------------------------------------------------------------ |
| [`typescript`](https://www.typescriptlang.org/)            | `^6.0.3`  | Language. Compiled to CommonJS into `dist/`.                                                     |
| [`ts-node`](https://typestrong.org/ts-node/)               | `^10.9.2` | Run `.ts` directly (`npm run cli` / `npx ts-node src/cli.ts`).                                   |
| [`@types/node`](https://www.npmjs.com/package/@types/node) | `^25.6.0` | Type definitions for Node core APIs.                                                             |
| [`playwright`](https://playwright.dev/)                    | `^1.59.1` | Drives the headed Chrome window during `login` and captures the outbound Bearer token + cookies. |
| [`@playwright/test`](https://playwright.dev/docs/intro)    | `^1.59.1` | Test-runner companion (present as a dev-dep; no live browser tests run in CI).                   |
| [`vitest`](https://vitest.dev/)                            | `^4.1.4`  | Test framework for the 208 unit + integration tests in `test_scripts/`.                          |

### External binaries you provide

- **Google Chrome** or **Microsoft Edge** (see Browser section above).
- **Node.js 20+** runtime.
- **git** for cloning.

No global npm tools need to be installed ahead of time. `npm install` in the
repo root fetches everything else.

---

## Build

```bash
git clone <this-repo> outlook-tool
cd outlook-tool
npm install
npm run build          # emits dist/cli.js (chmod +x in postbuild)
```

Optional: link the CLI globally so you can call it as `outlook-cli` from anywhere:

```bash
npm link                # installs a symlink at $(npm prefix -g)/bin/outlook-cli
```

If you previously linked a differently named package, remove the stale symlink
first (`rm $(which outlook-cli)`), then re-run `npm link`.

You can also run the TypeScript sources directly without building:

```bash
npx ts-node src/cli.ts <subcommand> [options]
# or
npm run cli -- <subcommand> [options]
```

---

## Run the tests

```bash
npm test               # vitest run — 388 tests across 34 files
npm run test:watch     # incremental
```

---

## First use

```bash
outlook-cli login
```

A Chrome window opens at `https://outlook.office.com/`. Sign in normally.
The tool watches outbound requests, captures the first Bearer token it sees,
closes the window, and writes `~/.outlook-cli/session.json`.

After that, every subcommand reads that file. You don't need to run `login`
again until the token expires (typically hours), and even then the default
behavior is to auto-reopen the browser for a silent refresh.

Quick verification:

```bash
outlook-cli auth-check
# {
#   "status": "ok",
#   "tokenExpiresAt": "2026-04-22T15:03:25.000Z",
#   "account": { "upn": "you@yourtenant.com" }
# }
```

---

## Configuration

Three runtime-plumbing settings exist. Each has a default, so **no
configuration is required** for a basic install:

| Setting                        | CLI flag                  | Env var                        | Default          |
| ------------------------------ | ------------------------- | ------------------------------ | ---------------- |
| Per-REST-call HTTP timeout     | `--timeout <ms>`          | `OUTLOOK_CLI_HTTP_TIMEOUT_MS`  | `30000` (30 s)   |
| Max wait for interactive login | `--login-timeout <ms>`    | `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` | `300000` (5 min) |
| Playwright Chrome channel      | `--chrome-channel <name>` | `OUTLOOK_CLI_CHROME_CHANNEL`   | `chrome`         |

Precedence: CLI flag > env var > default. A malformed flag or env value still
throws `ConfigurationError` (exit 3); the default only covers the unset case.

If you want persistent overrides, `source ./outlook-cli.env` in your shell or
append that line to `~/.zshrc` / `~/.bashrc`.

Other (always-optional) flags are listed in `outlook-cli --help` and in
`docs/design/configuration-guide.md`.

### Exit codes

| Code | Meaning                                                                                           |
| ---- | ------------------------------------------------------------------------------------------------- |
| `0`  | Success                                                                                           |
| `1`  | Unexpected error                                                                                  |
| `2`  | Invalid usage / bad argv                                                                          |
| `3`  | Configuration error (malformed flag or env var)                                                   |
| `4`  | Auth failure (expired/rejected session, user cancelled login, `--no-auto-reauth` with no cache)   |
| `5`  | Upstream API error (non-401 HTTP error, timeout, network failure, pagination limit)               |
| `6`  | IO error — includes folder collision without `--idempotent`, file collision without `--overwrite` |

---

## Usage examples

### Mail

```bash
# Most-recent 5 inbox messages as a human-readable table
outlook-cli list-mail --top 5 --table

# A specific message, body as text, written to disk
outlook-cli get-mail AAMkAGI... --body text > message.json

# Save all non-inline attachments to ./att
outlook-cli download-attachments AAMkAGI... --out ./att

# Incremental sync: every message received in the last 24h, paginated
outlook-cli list-mail \
  --folder Inbox \
  --since "$(date -u -v-24H +%Y-%m-%dT%H:%M:%SZ)" \
  --all --max 5000 --json

# A specific date range, all pages
outlook-cli list-mail \
  --since 2026-04-01T00:00:00Z \
  --until 2026-04-08T00:00:00Z \
  --all --json
```

`--since`/`--until` add a server-side `$filter` on `ReceivedDateTime`.
`--all` walks `@odata.nextLink` until exhausted; `--max <N>` is the
safety cap (default 10000, max 100000). When the cap is hit and more
results remain, a `max_results_reached` warning is emitted on stderr and
the partial result is returned.

### Send, reply, forward

```bash
# Compose a new email (opens as draft in Outlook desktop)
outlook-cli send-mail \
  --to "alice@example.com" "bob@example.com" \
  --cc "carol@example.com" \
  --subject "Q2 review" \
  --html body.html

# Send immediately (skip the draft)
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

# Reply-all (recipients pre-populated by M365)
outlook-cli reply-all AAMkAGI... --html reply.html --send-now

# Forward (--to is required)
outlook-cli forward AAMkAGI... \
  --to "dave@example.com" \
  --html note.html

# Extract your signature from the latest sent message
outlook-cli capture-signature
# or from a specific message
outlook-cli capture-signature --from-message AAMkAGI...
```

All send commands default to **draft-first** — the message is created as a
draft and Outlook desktop is activated so you can review before sending.
Pass `--send-now` to dispatch immediately. Automatic CC-self is on by default;
suppress with `--no-cc-self`. Signature from `~/.outlook-cli/signature.html`
is appended automatically; suppress with `--no-signature`.

### Calendar

```bash
# Next 14 days
outlook-cli list-calendar --from now --to "now + 14d" --table

# One event
outlook-cli get-event AAMkAGI...
```

### Folders

```bash
# List top-level folders
outlook-cli list-folders --table

# Walk the whole tree (bounded)
outlook-cli list-folders --recursive --table

# Resolve a folder by display-name path
outlook-cli find-folder "Inbox/Projects/Alpha"

# Create a nested folder idempotently (no-op if it already exists)
outlook-cli create-folder "Inbox/Projects/Alpha" --create-parents --idempotent

# Move messages to a folder (by alias, path, or id:<raw>)
outlook-cli move-mail AAMk... AAMk... --to "Inbox/Archive-2026"
outlook-cli move-mail AAMk... --to Archive
outlook-cli move-mail AAMk... --to "id:AAMkAGI..." --continue-on-error
```

### SharePoint reference attachments

Some Outlook messages include `ReferenceAttachment` entries — these are
SharePoint or OneDrive-for-Business shared-link "attachments" rather
than inline binaries. Their content lives on a different host
(`<tenant>.sharepoint.com`), which uses different cookies and a
different Bearer token from outlook.office.com.

To fetch them, capture a SharePoint session at login time:

```bash
# Capture both Outlook and SharePoint sessions in one login flow
outlook-cli login --sharepoint-host nbg.sharepoint.com
```

This writes a second session file to `~/.outlook-cli/sharepoint-session.json`
(mode 0600). Then download a referenced URL:

```bash
# Fetch a URL surfaced by get-mail's Attachments[].SourceUrl
outlook-cli download-sharepoint-link \
  "https://nbg.sharepoint.com/sites/foo/Documents/report.pdf" \
  --out ./att
```

If the SharePoint session file is missing or expired, the command exits
with code 4 (auth failure) and prints the exact `outlook-cli login`
command to recover.

### List mail from an arbitrary folder

```bash
# By display-name path (resolved once, then listed)
outlook-cli list-mail --folder "Inbox/Projects/Alpha" --top 10 --table

# By explicit id (skips resolution)
outlook-cli list-mail --folder-id AAMkAGI... --top 20

# With an anchor — resolve "Projects/Alpha" under Inbox
outlook-cli list-mail --folder-parent Inbox --folder "Projects/Alpha"
```

`outlook-cli <subcommand> --help` shows the complete flag set for each.

---

## Output modes

Every subcommand supports two formats:

- `--json` (default) — stable, stdout, pipe into `jq` / scripts.
- `--table` — human-readable, compact columns.

They are mutually exclusive. Errors are always emitted as JSON on **stderr**
with `code`, optional `message`, and setting-specific fields (e.g.
`missingSetting`, `destination`, `failed[]`).

---

## Project layout

```text
src/
  cli.ts                 # commander wiring, global options, error mapping
  auth/                  # Playwright login flow, token capture
  session/               # atomic session-file IO, locking, JWT parsing
  http/                  # OutlookClient + error types + REST DTOs
  folders/               # folder resolver (path, well-known, id) + types
  commands/              # one file per subcommand
  output/                # JSON / table formatter
  config/                # loadConfig, env + flag precedence, defaults
  util/                  # redaction, filename safety, misc helpers
test_scripts/            # vitest suites — 388 tests across 34 spec files
docs/
  design/                # refined specs, plans, project-design, config guide
  reference/             # codebase scans
  research/              # deep-dive docs on Outlook REST v2.0 quirks
  superpowers/           # workflow notes for repo automation
```

Every meaningful behavior is documented in
[`docs/design/project-design.md`](docs/design/project-design.md), and every
phase of work has a `plan-NNN-*.md` alongside it.

---

## Origin

Forked from [BikS2013/outlook-tool](https://github.com/BikS2013/outlook-tool)
by Giorgos Marinos, whose core insight — capturing an Outlook-web bearer
token via headed Playwright and reusing it against the `outlook.office.com/api/v2.0`
REST surface — made this approach viable.

The codebase has since been substantially rewritten and extended: folder
management, send/reply/forward with signature + inline-image support,
silent token renewal, atomic session storage with file locking, redaction
on every error path, and a 388-test vitest suite.

---

## License

MIT. See [LICENSE](LICENSE) for full text and dual copyright (original
upstream + this fork's substantial rewrite).
