# Codebase Scan: Outlook CLI

Scan date: 2026-04-21
Target request: `docs/design/refined-request-outlook-cli.md` (Outlook CLI tool with
Playwright-driven login, Bearer/cookie capture, and read-only REST commands).

---

## 1. Project Overview

- **Language / runtime**: TypeScript on Node.js (executed via `ts-node`; no build pipeline yet).
- **Module system**: CommonJS (`"type": "commonjs"` in `package.json`, `"module": "commonjs"` in `tsconfig.json`).
- **TS config**: `ES2022` target, `strict: true`, `esModuleInterop: true`, `skipLibCheck: true`, `types: ["node"]`, `outDir: dist`. Includes only `test_scripts/**/*.ts` — no `src/` yet.
- **Installed dependencies** (all `devDependencies`, verified in `node_modules/`):
  - `playwright@^1.59.1`
  - `@playwright/test@^1.59.1` (unused by the current script; test harness)
  - `@types/node@^25.6.0`
  - `ts-node@^10.9.2`
  - `typescript@^6.0.3`
- **Build/test scripts**: none defined beyond the placeholder `npm test`. No `bin` entry in `package.json`.
- **Directory layout** (root: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/`):
  - `CLAUDE.md` — project conventions (copy of the global Structure & Conventions).
  - `package.json`, `package-lock.json`, `tsconfig.json`.
  - `test_scripts/outlook_read_recent.ts` — the single existing implementation.
  - `docs/design/refined-request-outlook-cli.md` — the refined spec driving this project.
  - `docs/reference/` — reference material (contains `workflow-checkpoint.json`; this file is added here).
  - `docs/research/` — empty placeholder.
  - `prompts/` — empty placeholder.
  - `.playwright-profile/` — Chromium persistent profile already populated from an earlier headed login (`Default/`, `Local State`, etc.).
  - `.playwright-cli/` — Playwright CLI trace / console logs from interactive debugging sessions (incidental, not part of the product).
  - `.playwright/` — empty directory left over from Playwright tooling.
  - `.claude/` — Claude Code harness metadata.
  - `outlook_report.json` — sample output produced by the baseline script.
  - `node_modules/`.
- **Not present yet**: `src/`, `Issues - Pending Items.md`, `docs/design/project-design.md`, `docs/design/project-functions.MD`, `docs/design/configuration-guide.md`, any CLI binary, any test runner config.

## 2. Existing Artifacts

### 2.1 `test_scripts/outlook_read_recent.ts` (132 lines)

Baseline Playwright script. Behavior:

- Launches Chromium via `chromium.launchPersistentContext(PROFILE_DIR, { channel: 'chrome', headless: false, ... })`, where `PROFILE_DIR` is resolved to `../.playwright-profile/` (project root).
- Supports two flavors: `work` -> `https://outlook.office.com/mail/inbox`, `personal` -> `https://outlook.live.com/mail/0/inbox`.
- Waits up to 5 minutes for the URL to match `outlook.(office|live).com/mail` and the message list DOM (`div[role="option"]`) to appear.
- **Scrapes the DOM** (`page.evaluate`) for sender / subject / preview / received, with aria-label fallback.
- Writes a JSON array to `outlook_report.json` and echoes it between `---BEGIN_REPORT---` / `---END_REPORT---` sentinels.
- CLI args: positional `[flavor] [count] [outPath]`. No flags, no help, no exit codes beyond `0` / `1`.

**Relevance to the new CLI:**

- **Reuse patterns**:
  - Persistent Chrome profile via `launchPersistentContext` with `channel: 'chrome'` and `headless: false`.
  - Generous login wait loop + DOM sentinel for "inbox reached".
- **Supersede**:
  - DOM scraping — spec §5/§6 requires REST calls to `outlook.office.com/api/v2.0` with captured Bearer + cookies. DOM parsing goes away.
  - Mixed `work` / `personal` flavors — spec NG6 excludes `outlook.com` (personal). Only `outlook.office.com` remains.
  - Hard-coded timeouts (`LOGIN_TIMEOUT_MS = 5 * 60 * 1000`) — spec §8 makes login timeout mandatory config with no fallback.
  - `console.log` progress messages on stdout — spec §5 requires JSON on stdout and progress on stderr (with `--quiet` support).
  - Output file at `outlook_report.json` — replaced by stdout JSON, per-command schemas.
- **Leave alone**:
  - The script itself can stay in `test_scripts/` as a historical/reference artifact, or be removed once the new CLI ships. It is _not_ imported by anything.
- **Key reference lines**:
  - `test_scripts/outlook_read_recent.ts:15` — profile dir resolution.
  - `test_scripts/outlook_read_recent.ts:102-106` — persistent-context launch (exact pattern the new `auth/browser.ts` should adopt).
  - `test_scripts/outlook_read_recent.ts:30-47` — inbox-reached detection loop.

### 2.2 `.playwright-profile/` at project root

- Contains a fully populated Chromium profile (signed-in session from prior manual test).
- Per spec §6.3 / §7.1, the new CLI's profile must live at `$HOME/.outlook-cli/playwright-profile/` (mode `0700`) with an optional override via `OUTLOOK_CLI_PROFILE_DIR` / `--profile-dir`.
- **Recommendation**: deprecate the root `.playwright-profile/`. The new CLI should default to the `$HOME/...` location. Add the root directory to `.gitignore` (if/when a repo is created) and optionally delete it before release; keeping it now does no harm but it is not reused.

### 2.3 `outlook_report.json`

- Sample output of the baseline. Disposable. Not consumed by anything.

### 2.4 `.playwright/` and `.playwright-cli/`

- Residue from Playwright CLI / `npx playwright codegen` sessions. Not part of the product. Can be ignored (and `.gitignore`d).

### 2.5 `CLAUDE.md`

- Only contains the generic Structure & Conventions block. No `<toolName>` tool documentation registered yet.
- The new CLI must add its own `<outlook-cli>` block with one child entry per subcommand (spec AC-CLAUDEMD-UPDATED).

## 3. Conventions (from `CLAUDE.md` + current code)

- **TypeScript only** for any tool/script created in-project.
- **Strict TS** (`strict: true`) — the new code must compile clean with strict null checks.
- **Tests live under `test_scripts/`** (create if missing; already exists here).
- **Plans** under `docs/design/plan-xxx-<desc>.md`; global design in `docs/design/project-design.md`; functional requirements in `docs/design/project-functions.MD`; reference material in `docs/reference/`.
- **Prompts** under `prompts/` with sequential-number prefix.
- **No fallback defaults for mandatory configuration** — missing mandatory settings must raise a typed error (spec §8 defines which settings are mandatory vs. have explicit defaults). Any proposed exception must be written to the memory file _before_ implementation.
- **Tool docs** in `CLAUDE.md` use the XML shape `<toolName><objective/><command/><info/></toolName>`.
- **Issues tracker** at project root: `Issues - Pending Items.md` (pending first, ranked; completed below). Not created yet — will need to be added as soon as the first discrepancy is logged.
- **Import style**: CommonJS (`import { chromium } from 'playwright'` works because of `esModuleInterop`). Stick with the existing pattern.
- **Locating code** in future answers: always include folder + file + class/function + line number.
- **No VCS operations** unless explicitly requested. (There is currently no `.git` at project root anyway.)

## 4. Integration Points

### 4.1 Proposed source layout

No `src/` exists yet. Propose creating one and switching `tsconfig.json#include` to cover both `src/**/*.ts` and `test_scripts/**/*.ts`:

```
src/
  cli.ts                    # argv parser + command dispatch (bin entry)
  commands/
    login.ts
    authCheck.ts
    listMail.ts
    getMail.ts
    downloadAttachments.ts
    listCalendar.ts
    getEvent.ts
  auth/
    browser.ts              # launchPersistentContext + headed Chrome
    fetchHook.ts             # init script that hooks window.fetch
    tokenCapture.ts         # exposeBinding channel + JWT parsing
    session.ts              # schema, atomic write (0600), read, validation
  http/
    client.ts               # fetch wrapper with timeout + 401 retry-once
    headers.ts              # Authorization, X-AnchorMailbox, cookie serialization
  config/
    config.ts               # precedence resolver + ConfigurationError (no fallback)
    errors.ts               # typed error classes (ConfigurationError, AuthError, UpstreamError, IoError)
  output/
    json.ts
    table.ts
  util/
    paths.ts                # $HOME resolution, mkdir 0700, write 0600
    logger.ts               # stderr progress, --quiet, --log-file
test_scripts/
  ac-*.ts                   # one script per Acceptance Criterion from spec §9
```

- `package.json` additions: `"bin": { "outlook-cli": "dist/cli.js" }`, a `build` script (`tsc`), and a `start` / `dev` script via `ts-node src/cli.ts`.
- `tsconfig.json`: widen `include` to `["src/**/*.ts", "test_scripts/**/*.ts"]`; keep `outDir: dist`.

### 4.2 Playwright profile location

- Spec §6.3 / §7.1: default profile dir = `$HOME/.outlook-cli/playwright-profile/` (mode `0700`), overridable by `OUTLOOK_CLI_PROFILE_DIR` / `--profile-dir`.
- **Decision**: move away from the project-root `.playwright-profile/`. The new CLI will not read or write the old location. The old folder is retained on disk but unused; can be safely deleted.

### 4.3 Session file location

- Default `$HOME/.outlook-cli/session.json` (mode `0600`), overridable via `OUTLOOK_CLI_SESSION_FILE` / `--session-file`. Schema in spec §7.2. Atomic write (temp + rename) per §7.3.

### 4.4 CLAUDE.md tool registration

- New top-level `<outlook-cli>` block plus one nested entry per subcommand (`login`, `auth-check`, `list-mail`, `get-mail`, `download-attachments`, `list-calendar`, `get-event`). Required by AC-CLAUDEMD-UPDATED.

## 5. Gaps to Fill

Missing dev/runtime tooling the new CLI will need:

| Gap                                                        | Needed for                                                                                                          | Recommendation                                                                                                                                                                  |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| CLI argument parser                                        | Global flags + 7 subcommands with options                                                                           | Add `commander` (widely used, tree-shake-friendly, first-class TS types). `yargs` is an acceptable alternative but heavier.                                                     |
| JWT decoding                                               | Extract `exp`, `puid`, `tid` from the captured Bearer                                                               | Add `jwt-decode` (small, no crypto; we are not verifying signatures, only reading claims per spec §6.3 step 5).                                                                 |
| Table output                                               | `--table` mode for `list-mail` / `list-calendar`                                                                    | Add `cli-table3` (lightweight, TS-friendly) or hand-roll a minimal formatter to avoid an extra dep. Pick one in the design phase.                                               |
| Test runner                                                | Automated coverage for AC scripts                                                                                   | Consider `vitest` or `node --test`. Given existing `@playwright/test` is unused, either adopt it for end-to-end tests or add `vitest` for unit tests. Decide in the plan phase. |
| `bin` + build                                              | Ship `outlook-cli` as an executable                                                                                 | Add `"bin"` in `package.json`, a `build` script (`tsc`), and a shebang (`#!/usr/bin/env node`) in `src/cli.ts`.                                                                 |
| `.gitignore`                                               | Exclude `node_modules/`, `.playwright-profile/`, `.playwright-cli/`, `.playwright/`, `dist/`, `outlook_report.json` | Add before the first commit.                                                                                                                                                    |
| `Issues - Pending Items.md`                                | Required by conventions                                                                                             | Create at root on first tracked issue.                                                                                                                                          |
| `docs/design/project-design.md` and `project-functions.MD` | Required by conventions                                                                                             | Create during the design phase.                                                                                                                                                 |
| Typed error hierarchy                                      | Enforce spec §8 "no fallback" rule and exit codes 2/3/4/5/6                                                         | Implement in `src/config/errors.ts`; surface `ConfigurationError` for missing mandatory config.                                                                                 |
| Cookie-jar serialization                                   | Build `Cookie:` header from stored Playwright cookies on every REST call                                            | Write a small helper in `src/http/headers.ts`; no extra dep required.                                                                                                           |

No Microsoft Graph SDK or MSAL library is needed — spec NG4 / NG5 explicitly restrict the tool to `outlook.office.com/api/v2.0` REST + `fetch`-intercepted Bearer. Avoid pulling `@azure/msal-node` / `@microsoft/microsoft-graph-client`.

---

## Summary of findings

- Project is a thin TS + Playwright sandbox with one DOM-scraping script (`test_scripts/outlook_read_recent.ts`) and a pre-populated `.playwright-profile/` at the project root.
- The new CLI will **keep** the persistent-Chrome pattern but **replace** DOM scraping with REST calls against `outlook.office.com/api/v2.0`, gated by a captured Bearer token + cookie jar.
- A `src/` tree must be created; `tsconfig.json#include` must be widened; `package.json` needs `bin` + `build`.
- The `.playwright-profile/` at the project root will be deprecated in favor of `$HOME/.outlook-cli/playwright-profile/` per spec §6.3 / §7.1.
- Missing dev deps to add: `commander` (argv), `jwt-decode` (claim extraction), and a table formatter (`cli-table3` or hand-rolled). A test runner decision (`vitest` vs. `@playwright/test`) is pending for the design phase.
- Conventions artifacts not yet in place: `Issues - Pending Items.md`, `docs/design/project-design.md`, `docs/design/project-functions.MD`, `CLAUDE.md` tool-registration block. All will be created as the project progresses.

Absolute output path: `/Users/giorgosmarinos/aiwork/coding-platform/outlook-tool/docs/reference/codebase-scan-outlook-cli.md`
