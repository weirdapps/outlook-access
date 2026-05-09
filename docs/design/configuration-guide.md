# Configuration Guide — outlook-cli

## Configuration sources and precedence

The CLI resolves every setting through a fixed precedence chain. **Highest wins.**

1. **CLI flag** — e.g. `--timeout 30000`, passed on the command line for one invocation.
2. **Environment variable** — e.g. `OUTLOOK_CLI_HTTP_TIMEOUT_MS=30000` exported in the shell (or sourced from `outlook-cli.env`).
3. **Default** — allowed _only_ for the three runtime-plumbing settings listed below (`httpTimeoutMs`, `loginTimeoutMs`, `chromeChannel`), per the project-specific exception recorded in CLAUDE.md (2026-04-21). All other settings that were marked mandatory still throw `ConfigurationError` (exit code `3`) on miss — there is no silent fallback for those.

A `.env`-style file (like `outlook-cli.env` in the project root) is simply a convenient way to populate **source #2** — it does not introduce a new precedence level.

## Configuration parameters

### Runtime plumbing (defaults allowed — CLAUDE.md exception 2026-04-21)

CLI flag overrides env var overrides default. A malformed value from flag or env (e.g. non-integer, non-positive) still throws `ConfigurationError` — the default only covers the unset case.

| Name           | CLI flag                  | Env var                        | Default          | Purpose                                                                                                                                                |
| -------------- | ------------------------- | ------------------------------ | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| HTTP timeout   | `--timeout <ms>`          | `OUTLOOK_CLI_HTTP_TIMEOUT_MS`  | `30000` (30 s)   | Abort a single REST call to `outlook.office.com` after this many milliseconds. Positive integer.                                                       |
| Login timeout  | `--login-timeout <ms>`    | `OUTLOOK_CLI_LOGIN_TIMEOUT_MS` | `300000` (5 min) | Max time to wait for the user to finish logging in inside the Playwright-controlled Chrome window. Positive integer.                                   |
| Chrome channel | `--chrome-channel <name>` | `OUTLOOK_CLI_CHROME_CHANNEL`   | `chrome`         | Which Chrome or Edge build Playwright should launch. One of `chrome`, `chrome-beta`, `chrome-dev`, `msedge`, `msedge-beta`. Must be installed locally. |

### Optional (defaults allowed by the spec)

| Name              | CLI flag                | Env var                    | Default                                                              | Purpose                                                                              |
| ----------------- | ----------------------- | -------------------------- | -------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| Session file path | `--session-file <path>` | `OUTLOOK_CLI_SESSION_FILE` | `$HOME/.outlook-cli/session.json`                                    | Where captured cookies + Bearer token are persisted. Mode `0600`.                    |
| Profile directory | `--profile-dir <path>`  | `OUTLOOK_CLI_PROFILE_DIR`  | `$HOME/.outlook-cli/playwright-profile`                              | Playwright persistent profile so you don't re-login every browser open. Mode `0700`. |
| Timezone          | `--tz <iana>`           | —                          | System timezone (`Intl.DateTimeFormat().resolvedOptions().timeZone`) | IANA zone used for `list-calendar` window calculations.                              |
| Output mode       | `--json` / `--table`    | —                          | `--json`                                                             | JSON (scriptable) vs. human-readable table.                                          |
| Quiet             | `--quiet`               | —                          | `false`                                                              | Suppress progress messages on stderr.                                                |
| No auto reauth    | `--no-auto-reauth`      | —                          | `false`                                                              | On 401 / expired session, FAIL rather than re-opening the browser.                   |

### Command-local options

Several commands have their own options (`-n / --top`, `--folder`, `--from`, `--to`, `--body`, `--out`, `--overwrite`, `--include-inline`). See `outlook-cli <command> --help` for each.

## Recommended storage / management

- **Non-secret values** (timeouts, Chrome channel) can be tracked in `outlook-cli.env` at the project root. This file is safe to commit if you want team defaults; or add it to `.gitignore` and keep it local.
- **Secrets** (captured Bearer token, session cookies) are NEVER placed in config files. They live only in `$HOME/.outlook-cli/session.json` at mode `0600`, written atomically by the CLI itself.
- **Permanent setup**: append `source /abs/path/to/outlook-cli.env` to `~/.zshrc` or `~/.bashrc` so every shell has the three mandatory settings preloaded.
- **One-off invocations**: prefix the command, e.g.

  ```bash
  OUTLOOK_CLI_HTTP_TIMEOUT_MS=30000 \
  OUTLOOK_CLI_LOGIN_TIMEOUT_MS=300000 \
  OUTLOOK_CLI_CHROME_CHANNEL=chrome \
    node dist/cli.js list-mail -n 5
  ```

- **Per-call override**: the CLI flag always wins, so you can bump a single call's timeout without touching env:

  ```bash
  node dist/cli.js --timeout 60000 list-mail -n 50
  ```

## Expiring values — proposal

None of the three mandatory settings expire. They are static runtime knobs.

The **session file** (`$HOME/.outlook-cli/session.json`) contains a short-lived Bearer token whose `bearer.expiresAt` field is an ISO-8601 timestamp derived from the JWT `exp` claim. The CLI already:

- Checks `expiresAt` before every call (`auth-check` reports `status: "expired"`).
- Triggers a browser re-auth automatically on expiry (unless `--no-auto-reauth`).
- Rotates the token transparently on 401.

If future iterations add settings that expire (e.g., a long-lived personal access token for an alternate auth backend), follow the CLAUDE.md guidance and add a companion setting to capture the expiration date, so the CLI can proactively warn the user before expiry.

## Validation errors

If a mandatory setting is missing, stderr receives pretty-printed JSON and the process exits 3. Example:

```json
{
  "error": {
    "code": "CONFIG_MISSING",
    "message": "Mandatory setting \"httpTimeoutMs\" was not provided. Checked: --timeout flag, OUTLOOK_CLI_HTTP_TIMEOUT_MS env var.",
    "missingSetting": "httpTimeoutMs",
    "checkedSources": ["--timeout flag", "OUTLOOK_CLI_HTTP_TIMEOUT_MS env var"]
  }
}
```

Numeric settings also validate ranges — a non-positive integer for either timeout will surface a `ConfigurationError` with a range message.
