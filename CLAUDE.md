<structure-and-conventions>
## Structure & Conventions

- Every time you want to create a test script, you must create it in the test_scripts folder. If the folder doesn't exist, you must make it.

- All the plans must be kept under the docs/design folder inside the project's folder in separate files: Each plan file must be named according to the following pattern: plan-xxx-<indicative description>.md

- The complete project design must be maintained inside a file named docs/design/project-design.md under the project's folder. The file must be updated with each new design or design change.

- All the reference material used for the project must be collected and kept under the docs/reference folder.
- All the functional requirements and all the feature descriptions must be registered in the /docs/design/project-functions.MD document under the project's folder.

<configuration-guide>
- If the user ask you to create a configuration guide, you must create it under the docs/design folder, name it configuration-guide.md and be sure to explain the following:
  - if multiple configuration options exist (like config file, env variables, cli params, etc) you must explain the options and what is the priority of each one.
  - Which is the purpose and the use of each configuration variable
  - How the user can obtain such a configuration variable
  - What is the recomented approach of storing or managing this configuration variable
  - Which options exist for the variable and what each option means for the project
  - If there are any default value for the parameter you must present it.
  - For configuration parameters that expire (e.g., PAT keys, tokens), I want you to propose to the user adding a parameter to capture the parameter's expiration date, so the app or service can proactively warn users to renew.
</configuration-guide>

- Every time you create a prompt working in a project, the prompt must be placed inside a dedicated folder named prompts. If the folder doesn't exists you must create it. The prompt file name must have an sequential number prefix and must be representative to the prompt use and purpose.

- You must maintain a document at the root level of the project, named "Issues - Pending Items.md," where you must register any issue, pending item, inconsistency, or discrepancy you detect. Every time you fix a defect or an issue, you must check this file to see if there is an item to remove.
- The "Issues - Pending Items.md" content must be organized with the pending items on top and the completed items after. From the pending items the most critical and important must be first followed by the rest.

- When I ask you to create tools in the context of a project everything must be in Typescript.
- Every tool you develop must be documented in the project's Claude.md file
- The documentation must be in the following format:
<toolName>
    <objective>
        what the tool does
    </objective>
    <command>
        the exact command to run
    </command>
    <info>
        detailed description of the tool
        command line parameters and their description
        examples of usage
    </info>
</toolName>

- Every time I ask you to do something that requires the creation of a code script, I want you to examine the tools already implemented in the scope of the project to detect if the code you plan to write, fits to the scope of the tool.
- If so, I want you to implement the code as an extension of the tool, otherwise I want you to build a generic and abstract version of the code as a tool, which will be part of the toolset of the project.
- Our goal is, while the project progressing, to develop the tools needed to test, evaluate, generate data, collect information, etc and reuse them in a consistent manner.
- All these tools must be documented inside the CLAUDE.md to allow their consistent reuse.

- When I ask you to locate code, I need to give me the folder, the file name, the class, and the line number together with the code extract.
- Don't perform any version control operation unless I explicitly request it.

- When you design databases you must align with the following table naming conventions:
  - Table names must be singular e.g. the table that keeps customers' data must be called "Customer"
  - Tables that are used to express references from one entity to another can by plural if the first entity is linked to many other entities.
  - So we have "Customer" and "Transaction" tables, we have CustomerTransactions.

- You must never create fallback solutions for configuration settings. In every case a configuration setting is not provided you must raise the appropriate exception. You must never substitute the missing config value with a default or a fallback value.
- If I ask you to make an exception to the configuration setting rule, you must write this exception in the projects memory file, before you implement it.
</structure-and-conventions>

## Project-specific exceptions to global rules

### Exception — defaults allowed for three runtime-plumbing config settings

On 2026-04-21 the user explicitly asked me to introduce defaults for three
settings that the refined spec §8 had marked "mandatory, no default":

- `httpTimeoutMs` — default **30000** (30 s per REST call). Env:
  `OUTLOOK_CLI_HTTP_TIMEOUT_MS`. Flag: `--timeout`.
- `loginTimeoutMs` — default **300000** (5 min for interactive login). Env:
  `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`. Flag: `--login-timeout`.
- `chromeChannel` — default **`"chrome"`**. Env:
  `OUTLOOK_CLI_CHROME_CHANNEL`. Flag: `--chrome-channel`.

Rationale: these three values are operational plumbing, not secrets or
environment-distinguishing identities, so forcing the user to set them on
every invocation (or in every shell) trades ergonomics for safety the
rule was designed to protect. The user accepted this trade-off.

Precedence remains unchanged — CLI flag > env var > **default** (new tier).
`loadConfig()` no longer throws `CONFIG_MISSING` for these three. Every
other mandatory setting (none today, but if future ones are added) must
continue to follow the global no-fallback rule unless a similar exception
is recorded here.

Implementation landed in `src/config/config.ts` (`DEFAULTS` constant +
`resolveOptionalInt` / `resolveOptionalString` helpers).

## Tools

<outlook-cli>
    <objective>
        CLI tool to authenticate against Outlook web (outlook.office.com), capture
        session cookies + bearer token via a headed Playwright Chrome browser,
        persist them safely under `$HOME/.outlook-cli/`, and access inbox,
        calendar, and attachments through the Outlook REST v2.0 API
        (`https://outlook.office.com/api/v2.0/...`). All subcommands read the
        cached session; expired or rejected sessions trigger a single
        automatic browser re-auth unless `--no-auto-reauth` is passed.
    </objective>
    <command>
        npx ts-node src/cli.ts <subcommand> [options]
        # after `npm run build`:
        node dist/cli.js <subcommand> [options]
        # or via npm run:
        npm run cli -- <subcommand> [options]
    </command>
    <info>
        Global flags (apply to every subcommand):
        - `--timeout <ms>`           Per-REST-call HTTP timeout (env
                                     `OUTLOOK_CLI_HTTP_TIMEOUT_MS`). Default: 30000.
        - `--login-timeout <ms>`     Max wait for interactive login (env
                                     `OUTLOOK_CLI_LOGIN_TIMEOUT_MS`). Default: 300000.
        - `--chrome-channel <name>`  Playwright Chrome channel (env
                                     `OUTLOOK_CLI_CHROME_CHANNEL`). Default: `chrome`.
                                     Other examples: `chrome-beta`, `msedge`.
        - `--session-file <path>`    Override session file. Default:
                                     `$HOME/.outlook-cli/session.json` (mode 0600).
        - `--profile-dir <path>`     Override Playwright profile dir. Default:
                                     `$HOME/.outlook-cli/playwright-profile/` (mode 0700).
        - `--tz <iana>`              IANA timezone. Defaults to system TZ.
        - `--json`                   Emit JSON (default).
        - `--table`                  Emit human-readable table (mutually exclusive with `--json`).
        - `--quiet`                  Suppress stderr progress messages.
        - `--no-auto-reauth`         On 401 or expired session, FAIL instead of re-opening browser.
        - `--log-file <path>`        Write debug log to file (mode 0600).

        Exit codes:
          0 success
          2 invalid argv / usage
          3 configuration error (missing mandatory config)
          4 auth failure (user cancellation, login timeout, 401-after-retry, or
            --no-auto-reauth with missing/expired session)
          5 upstream API error (non-401 HTTP error, timeout, network failure)
          6 IO error (cannot write session file, dir permission, file collision
            without --overwrite)
          1 unexpected error

        Subcommands:

        1. `login [--force]`
           - Opens Chrome via Playwright, waits for the user to log into Outlook,
             captures the first Bearer token + cookies, writes the session file.
           - With `--force`, always opens the browser (no cache reuse).
           - Without `--force`, returns the cached session directly if it exists
             and is not expired.
           - Output: `{status, sessionFile, tokenExpiresAt, account:{upn, puid, tenantId}}`.

        2. `auth-check`
           - Loads the session and calls `GET /api/v2.0/me` with
             `noAutoReauth: true` to verify the token is still accepted.
           - Never opens the browser. Always exits 0 (the status is reported
             in the payload).
           - Output: `{status: "ok"|"expired"|"missing"|"rejected", tokenExpiresAt, account}`.

        3. `list-mail [-n <N>] [--folder <name>] [--folder-id <id>] [--folder-parent <anchor>] [--select <csv>]`
           - Lists recent messages from a folder (well-known alias, display-name
             path, or raw id).
           - `--top N`           1..100 (default 10).
           - `--folder`          One of `Inbox`, `SentItems`, `Drafts`,
                                 `DeletedItems`, `Archive` (original fast path,
                                 no resolver hop) OR any other well-known alias
                                 (`JunkEmail`, `Outbox`, `MsgFolderRoot`,
                                 `RecoverableItemsDeletions`) OR a display-name
                                 path (e.g. `Inbox/Projects/Alpha`). Default:
                                 `Inbox`.
           - `--folder-id <id>`  Raw folder id. XOR with `--folder` — passing
                                 both → exit 2. When set, the resolver is
                                 bypassed and the id is used verbatim.
           - `--folder-parent`   Anchor folder (well-known alias, path, or
                                 `id:<raw>`) used when `--folder` is a bare
                                 name / path. Default `MsgFolderRoot`. Illegal
                                 with `--folder-id` or alone (without
                                 `--folder`) → exit 2.
           - `--select`          Comma-separated $select fields. Default:
                                 `Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead,WebLink`.
           - JSON: array of `MessageSummary`.
           - Table columns: `Received | From | Subject | Att | Id`.

        4. `get-mail <id> [--body <html|text|none>]`
           - Retrieves one message. `id` is positional and required.
           - `--body`     `html` (raw HTML Body passed through),
                          `text` (default; upstream Body passed through untouched
                           — HTML→text conversion is deferred),
                          `none` (omit the Body field).
           - Fetches `/api/v2.0/me/messages/{id}` plus `.../attachments` metadata
             and merges them as `Attachments: AttachmentSummary[]` on the result.

        5. `download-attachments <id> --out <dir> [--overwrite] [--include-inline]`
           - Saves FileAttachment content bytes into `--out` (created with mode
             0700 if missing).
           - `--out` is mandatory; missing → exit 3.
           - Skips inline attachments unless `--include-inline` is set.
           - Skips ReferenceAttachment and ItemAttachment (recorded in
             `skipped[]` with the appropriate `reason`).
           - Without `--overwrite`, colliding filenames exit 6; duplicate names
             within the same run are auto-suffixed `" (1)"`, `" (2)"`, ...
           - Output: `{messageId, outDir, saved:[{id,name,path,size}],
                      skipped:[{id,name,reason,sourceUrl?,odataType?}]}`.

        6. `list-calendar [--from <ISO>] [--to <ISO>] [--tz <iana>]`
           - `--from` accepts ISO8601, `now`, or `now + Nd` (default `now`).
           - `--to`   accepts ISO8601, `now`, or `now + Nd` (default `now + 7d`).
           - Calls `GET /api/v2.0/me/calendarview?startDateTime=...&endDateTime=...
             &$orderby=Start/DateTime asc&$select=Id,Subject,Start,End,
             Organizer,Location,IsAllDay`.
           - JSON: array of `EventSummary`.
           - Table columns: `Start | End | Subject | Organizer | Location | Id`.

        7. `get-event <id> [--body <html|text|none>]`
           - Retrieves a single event. Body handling identical to get-mail.

        8. `list-folders [--parent <spec>] [--top <N>] [--recursive] [--include-hidden] [--first-match]`
           - Enumerates mail folders under a parent.
           - `--parent <spec>`   Well-known alias, display-name path, or
                                 `id:<raw>`. Default `MsgFolderRoot`.
           - `--top N`           Per-page `$top` (1..250, default 100).
           - `--recursive`       Walk the full sub-tree (bounded by the
                                 internal 5000-folder safety cap). Materializes
                                 a `Path` field on each row (escaped
                                 slash-separated — `/` becomes `\/`, `\`
                                 becomes `\\`).
           - `--include-hidden`  Include folders whose `IsHidden === true`.
                                 Default: false.
           - `--first-match`     On ambiguity during `--parent` resolution,
                                 pick the oldest candidate (`CreatedDateTime`
                                 ascending, `Id` ascending) instead of exit 2.
           - JSON: array of `FolderSummary` objects with a materialized `Path`.
           - Table columns: `Path | Unread | Total | Children | Id`.

        9. `find-folder <spec> [--anchor <spec>] [--first-match]`
           - Resolves a folder query to a single `ResolvedFolder` (including
             the resolver's provenance in `ResolvedVia`).
           - `<spec>` (required) — one of:
               - a well-known alias (`Inbox`, `Archive`, …),
               - a display-name path (`Inbox/Projects/Alpha`),
               - `id:<raw>` for a direct GET on the opaque id.
           - `--anchor <spec>`   Anchor for path-form queries. Ignored for
                                 well-known / id queries. Default
                                 `MsgFolderRoot`.
           - `--first-match`     Tiebreaker on ambiguity (see `list-folders`).
           - Exit codes:
               - 5 `UPSTREAM_FOLDER_NOT_FOUND` — the folder or any path
                 segment does not exist.
               - 2 `FOLDER_AMBIGUOUS` — multiple siblings share the same
                 DisplayName (add `--first-match` or use `id:<raw>`).
           - JSON: single `ResolvedFolder` object with `ResolvedVia:
             "wellknown" | "path" | "id"`.

        10. `create-folder <path-or-name> [--parent <spec>] [--create-parents] [--idempotent]`
           - Creates a folder (or a nested path) under an anchor.
           - `<path-or-name>` (required) — a bare name (`Alpha`) or a
             slash-separated display-name path (`Projects/Alpha`). A well-known
             alias is rejected when the anchor is `MsgFolderRoot`. Escape
             rules: `/` inside a DisplayName is `\/`, `\` is `\\`.
           - `--parent <spec>`      Anchor folder (well-known, path, or
                                    `id:<raw>`). Default `MsgFolderRoot`.
           - `--create-parents`     Create missing intermediate segments.
                                    Without it, a missing intermediate →
                                    exit 2 `FOLDER_MISSING_PARENT`.
           - `--idempotent`         Treat a `FOLDER_ALREADY_EXISTS` collision
                                    (HTTP 400 or 409 with OData
                                    `error.code === 'ErrorFolderExists'`) as
                                    success and return the pre-existing folder
                                    (`PreExisting: true`, top-level
                                    `idempotent: true`). Without this flag,
                                    the collision exits 6 with
                                    `FOLDER_ALREADY_EXISTS`.
           - JSON: `CreateFolderResult` (`{ created:[…], leaf:…, idempotent:
             boolean }`) — `created[]` entries carry `Path`, `Id`,
             `ParentFolderId`, `PreExisting`.
           - Table columns: `Path | Id | PreExisting` (applied to
             `result.created[]`).

        11. `move-mail <messageIds...> --to <spec> [--first-match] [--continue-on-error]`
           - Moves one or more messages to a destination folder.
           - **IMPORTANT — move returns a NEW id.** Outlook's
             `POST /me/messages/{id}/move` responds with a new message
             identity in the destination folder; the source id is no longer
             resolvable. The command surfaces the pairing explicitly in
             `moved[]` so scripts don't chain stale ids.
           - `<messageIds...>` (required) — one or more source message ids.
           - `--to <spec>` (required) — destination folder: well-known alias,
             display-name path, or `id:<raw>`. Aliases are always pre-resolved
             to a raw id before the `/move` POST (ADR-16).
           - `--first-match`         Tiebreaker on ambiguity during `--to`
                                     resolution.
           - `--continue-on-error`   Collect per-message failures in
                                     `failed[]` instead of aborting. The
                                     process still exits 5 when `failed[]`
                                     is non-empty (partial-failure rule).
           - JSON: `MoveMailResult` with `destination`, `moved[]` (each entry
             `{ sourceId, newId }`), `failed[]` (each entry
             `{ sourceId, error:{ code, httpStatus?, message? } }`), and a
             `summary: { requested, moved, failed }`.
           - Table columns: `Source Id | New Id | Status | Error`.

        Folder error codes (additional to the generic upstream taxonomy):
          - `UPSTREAM_FOLDER_NOT_FOUND`   — exit 5 (folder or path segment
                                             absent).
          - `UPSTREAM_PAGINATION_LIMIT`   — exit 5 (50-page per-collection
                                             cap or 5000-node tree cap).
          - `FOLDER_PATH_INVALID`         — exit 2 (bad escape, empty
                                             segment, > 16 segments).
          - `FOLDER_MISSING_PARENT`       — exit 2 (intermediate segment
                                             absent without --create-parents).
          - `FOLDER_AMBIGUOUS`            — exit 2 (multiple siblings share
                                             a DisplayName; use --first-match
                                             or id:<raw>).
          - `FOLDER_ALREADY_EXISTS`       — exit 6 (leaf collision without
                                             --idempotent).

        Examples:

        First-time login (sets mandatory config via env for this shell):
        ```bash
        export OUTLOOK_CLI_HTTP_TIMEOUT_MS=30000
        export OUTLOOK_CLI_LOGIN_TIMEOUT_MS=300000
        export OUTLOOK_CLI_CHROME_CHANNEL=chrome
        npx ts-node src/cli.ts login
        ```

        Verify the session, list 5 most-recent inbox messages, download attachments:
        ```bash
        npx ts-node src/cli.ts auth-check
        npx ts-node src/cli.ts list-mail --top 5 --table
        npx ts-node src/cli.ts get-mail AAMkAGI... --body text > message.json
        npx ts-node src/cli.ts download-attachments AAMkAGI... --out ./att
        ```

        Calendar:
        ```bash
        npx ts-node src/cli.ts list-calendar --from now --to "now + 14d" --table
        npx ts-node src/cli.ts get-event AAMkAGI...
        ```

        Folders — enumerate, resolve, create, move, list-in:
        ```bash
        # Top-level folders; recursive walk with --table
        npx ts-node src/cli.ts list-folders --table
        npx ts-node src/cli.ts list-folders --parent Inbox --recursive --table

        # Resolve a path to an id
        npx ts-node src/cli.ts find-folder "Inbox/Projects/Alpha" --json

        # Create a path; idempotent re-run returns the pre-existing folder
        npx ts-node src/cli.ts create-folder "Projects/Alpha" --parent Inbox --create-parents
        npx ts-node src/cli.ts create-folder "Projects/Alpha" --parent Inbox --create-parents --idempotent

        # List messages in a user folder (by path or by id)
        npx ts-node src/cli.ts list-mail --folder "Inbox/Projects/Alpha" -n 5
        npx ts-node src/cli.ts list-mail --folder-id AAMkAGI... -n 5

        # Move messages (surface the new ids in moved[])
        npx ts-node src/cli.ts move-mail AAMkAGI...srcA AAMkAGI...srcB \
          --to "Inbox/Projects/Alpha" --continue-on-error
        ```

        Security notes:
        - The bearer token and cookie values are NEVER logged or printed.
        - The session file is written atomically (write + fsync + rename) with
          mode 0600 inside a 0700 parent directory.
        - A PID-based advisory lock at `$HOME/.outlook-cli/.browser.lock`
          prevents two concurrent login flows from racing on the profile dir.
    </info>
</outlook-cli>

