<!-- markdownlint-disable-next-line MD041 -->
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
- Every tool you develop must be documented in `README.md`, and linked from here.
  **Relaxed 2026-09-13**, from "documented in the project's CLAUDE.md" in the
  `<toolName><objective><command><info>` XML format. The rule was right about the
  requirement and wrong about the location: CLAUDE.md is re-read on every session
  in this repo, so a 317-line tool spec was charged to every one of them, and it
  had drifted into a near-complete duplicate of `README.md` rather than a source
  of truth. The XML shape is retained as the section skeleton in the README
  (objective, command, parameters with defaults and env vars, examples, exit
  codes, security notes); only the host file changed.
- Documentation must still cover, for each tool: what it does, the exact command,
  every parameter with its default and env-var override, worked examples, exit
  codes, and anything security-relevant. A flag that exists in `src/` and not in
  `README.md` is a documentation bug.

- Every time I ask you to do something that requires the creation of a code script, I want you to examine the tools already implemented in the scope of the project to detect if the code you plan to write, fits to the scope of the tool.
- If so, I want you to implement the code as an extension of the tool, otherwise I want you to build a generic and abstract version of the code as a tool, which will be part of the toolset of the project.
- Our goal is, while the project progressing, to develop the tools needed to test, evaluate, generate data, collect information, etc and reuse them in a consistent manner.
- All these tools must be documented in `README.md` to allow their consistent reuse, and linked from here so they stay discoverable.

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

`outlook-cli` is documented in full in [`README.md`](README.md): objective and
auth model under "Why this exists" and "How it works", every subcommand under
"Usage", global flags and the two output modes under "Output modes",
environment overrides under "Configuration", the `~/.outlook-cli/` runtime
layout and exit codes under "Configuration", and the token/lock guarantees
under "Security posture".

**Read it there, not here.** This block used to carry a 317-line copy of that
spec, which every session in this repo paid for and which had drifted into a
near-duplicate. Before it was removed on 2026-09-13 the README was checked to
carry all 48 of its flags; the four things it was missing (`--quiet`,
`--log-file`, the `.browser.lock` advisory lock and the fsync detail) were
written into the README first, each verified against `src/`.

Quick reference, because these three are the ones that bite:

- `outlook-cli login` when a tool returns `auth_required`. Exit code **4 always
  means re-authenticate**, and **5 always means upstream misbehaved**, across
  all three M365 CLIs.
- `--concurrency` above 2 gets you `ApplicationThrottled` (HTTP 429) from M365.
- The renewable credential is the `playwright-profile/` directory, not
  `session.json`. Copying the JSON alone gives you working tokens until they
  expire and then no renewal path.

## The three M365 CLIs

One surface per repo. They share no session, no Playwright profile and no
login, and each has its own state directory.

| Surface                     | Repo                | Invoke                                            | State dir            |
| --------------------------- | ------------------- | ------------------------------------------------- | -------------------- |
| Mail, calendar, attachments | `outlook-access`    | `outlook-cli` (on PATH)                           | `~/.outlook-cli/`    |
| Teams chats and channels    | `teams-access`      | `node ~/SourceCode/teams-access/dist/cli.js`      | `~/.teams-cli/`      |
| SharePoint + OneDrive files | `sharepoint-access` | `node ~/SourceCode/sharepoint-access/dist/cli.js` | `~/.sharepoint-cli/` |

Exit codes 0-6 are identical across all three: **4 = re-authenticate**,
**5 = upstream error**. Keep it that way; cron wrappers branch on those numbers.

### SharePoint moved out (2026-08-08)

SharePoint and OneDrive-for-Business now live in `sharepoint-access`. The code
still here is **deprecated and scheduled for removal**:
`src/auth/sharepoint-capture.ts`, `src/http/sharepoint-client.ts`,
`src/session/sharepoint-schema.ts`, the `download-sharepoint-link` command and
the `--sharepoint-host` flag on `login`/`auth-renew`.

Do not build anything new on them. The flag is still wired into
`sync-tokens-to-vps.sh` as a belt-and-braces second session producer while the
new path completes an observation cycle; see P9 in
`sharepoint-access/Issues - Pending Items.md` for the removal checklist.
