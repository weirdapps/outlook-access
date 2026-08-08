# sharepoint-access: A Standalone SharePoint CLI

**Date:** 2026-08-08
**Repos:** `sharepoint-access` (new), `outlook-access` (removal), `plessas-marketplace` (new `files` plugin), `teams-access` (unrelated fix, appendix A)
**Status:** Approved design, pending implementation plan
**Supersedes:** the first revision of this file, which put `sp-*` subcommands on `outlook-cli`

---

## 1. Problem

`outlook-access` captures a working SharePoint session but exposes almost nothing
on top of it. `SharepointClient` has exactly one method, `getBinary(absoluteUrl)`,
surfaced as one command, `download-sharepoint-link`. There is no way to browse a
document library, list a folder, download by path, search, create a folder, or
upload.

The auth layer is not the gap. Capture listens at `context.on('request')`, which
catches Service-Worker-dispatched and MCAS-proxied requests that a page-level or
in-page hook misses. Session persistence, the `.browser.lock` concurrency guard,
and headless silent renewal all work. The gap is purely API surface.

There is a second, structural problem. This ecosystem's established pattern is one
Microsoft surface per repo: `outlook-access` provides `outlook-cli` with
`~/.outlook-cli/` and its own Playwright profile, bridged by the `mail` plugin;
`teams-access` provides `teams-cli` with `~/.teams-cli/` and a completely separate
Playwright profile, bridged by the `chat` plugin. The two share no profile, no
session store, and no login. SharePoint is the sole exception, living inside
`outlook-access` because its SSO happened to be capturable from the Outlook
browser context.

Growing the SharePoint surface inside `outlook-access` would entrench that
exception. `outlook-access` is already the heavy repo at 41 files and 10,233
lines covering mail, calendar, folders, attachments, signatures, and threads,
against `teams-access` at 25 files and 2,684 lines.

## 2. Decision

Build `sharepoint-access`, providing `sharepoint-cli`, with its own
`~/.sharepoint-cli/` directory, its own Playwright profile, its own interactive
login, and its own MCP bridge plugin. Remove SharePoint from `outlook-access`
once the new repo works.

Three facts make now the right moment:

- SharePoint is only 537 lines across 4 files in `outlook-access` today, and this
  design adds roughly 1,500 more. Extracting costs 537 lines now or 2,000 later.
- Nothing outside `outlook-access` calls `--sharepoint-host` or
  `download-sharepoint-link`. `plessas-marketplace`, `atm-recon`, and
  `claude-config` were all searched. The only consumer is the re-auth line in the
  global `CLAUDE.md`, a one-line edit.
- The pattern to copy already exists and is proven twice.

The accepted cost: SharePoint SSO currently rides free on the Outlook login,
because `captureSharepointFromContext` reuses a context that has just
authenticated. A standalone repo needs its own profile and one interactive MFA at
setup, then silent renewal via its own `ESTSAUTHPERSISTENT` for roughly 90 days.
This is exactly the deal `teams-access` already makes.

### 2.1 Rejected alternatives

**Reuse `~/.outlook-cli/playwright-profile` from the new repo.** Would keep SSO
free and avoid any extra MFA, but couples two independently-versioned binaries to
a shared filesystem path and a shared lock. Two Chrome instances racing one
profile directory is a genuinely nasty failure mode, and `teams-access` already
demonstrates that a separate profile is acceptable.

**Extract a shared `ms-auth-core` package** for jwt decoding, locking, session
store, config loading, and output formatting. The textbook answer, and wrong
here: `teams-access` already duplicates all of it, so a third copy is consistent
with the ecosystem while a shared package would couple three release cycles for
roughly 500 lines. Revisit only if a fourth Microsoft surface appears.

## 3. Established constraints

Measured against the live tenant on 2026-08-08, not assumed.

| Probe                           | Result                                    | Consequence                              |
| ------------------------------- | ----------------------------------------- | ---------------------------------------- |
| Session contents                | `FedAuth` + `rtFa` cookies, **no Bearer** | Cookie auth is the only credential       |
| `GET /_api/web`                 | 200                                       | Classic SPO REST is available            |
| `GET /_api/v2.0/me/drive`       | 403 `accessDenied`                        | Graph-shaped `driveItem` API is blocked  |
| `GET /_api/v2.0/sites/root`     | 403 `accessDenied`                        | Same                                     |
| `POST /_api/contextinfo`        | 200, `FormDigestTimeoutSeconds: 1800`     | Writes are feasible, digest lives 30 min |
| `GET /_api/search/query`        | 200                                       | Search and site discovery are available  |
| `GetFolderByServerRelativeUrl`  | 200                                       | Legacy accessor works                    |
| `GetFolderByServerRelativePath` | 200                                       | Modern accessor works, and is preferred  |

Two conclusions follow and are not revisitable without new evidence:

1. **Microsoft Graph is unusable for SharePoint here.** `graph.microsoft.com`
   requires a Graph-audience Bearer. The tenant issues none for SharePoint, and
   cookies do not authorise Graph. The tenant additionally blocks the
   SharePoint-hosted `/_api/v2.0` mirror of the same object model.
2. **Writes need CSRF handling.** Cookie-authenticated POSTs to SPO REST are
   rejected without an `X-RequestDigest` header. Bearer callers are exempt, which
   is why the requirement is easy to miss when reading Graph-oriented docs.

## 4. Non-goals

- No shared auth package across the three repos. See 2.1.
- No long-lived local proxy. A resident browser plus an unauthenticated
  `127.0.0.1` port fronting a document store is worse than persist-and-renew,
  which is stateless between invocations, survives reboot, and suits cron.
- No OneDrive-for-Business personal-site support in this pass.
- No `outlook-cli` compatibility shim. The removal in phase 5 is a clean break.

## 5. Repository layout

Mirrors `teams-access`, which is the closest structural precedent.

```text
sharepoint-access/
  src/
    auth/
      capture.ts          standalone interactive + headless capture
      lock.ts             ported from outlook-access
    session/
      schema.ts           v1: host, bearer?, cookies, capturedAt, tokenExpiresAt
      store.ts            load/save ~/.sharepoint-cli/session.json
    config/
      load.ts  errors.ts
    http/
      sharepoint-client.ts  getJson, postJson, getBinary, putBinary
      errors.ts  types.ts
    sharepoint/
      paths.ts            server-relative path handling
      digest.ts           X-RequestDigest fetch + cache
      upload.ts           size-routed upload strategy
    commands/
      login.ts  auth-renew.ts  auth-check.ts  health-check.ts
      ls.ts  get.ts  put.ts  mkdir.ts  search.ts  libraries.ts
    output/
      json.ts  table.ts
    util/
      redact.ts  exit-codes.ts
    cli.ts
  test_scripts/*.spec.ts
  docs/design/{project-design.md, project-functions.MD}
  CLAUDE.md
  .pre-commit-config.yaml
```

Runtime state lives in `~/.sharepoint-cli/`: `playwright-profile/`,
`session.json`, `.browser.lock`.

Test files use the `.spec.ts` suffix, matching `outlook-access` (the source of
the extracted code) rather than `teams-access`, which uses `.test.ts`. The two
existing repos are already inconsistent on this; the new repo picks one and its
`vitest.config.ts` enforces it.

## 6. Authentication

Ported from `src/auth/sharepoint-capture.ts`, with one substantive change: it
must handle a **cold** profile. The current code assumes a warm context whose
Microsoft SSO cookies were just established by an Outlook sign-in, so it
navigates and expects silent completion. Standalone, the first `login` faces the
full interactive redirect chain including MFA.

| Command                                  | Behaviour                                                                                                                                                                                                                                                                      |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `login [--host <tenant>.sharepoint.com]` | Launch persistent Chrome on `~/.sharepoint-cli/playwright-profile`, navigate to `https://<host>/_layouts/15/sharepoint.aspx`, wait up to the interactive timeout for the user to complete sign-in, capture cookies (and a Bearer if the tenant emits one), persist the session |
| `auth-renew`                             | Headless relaunch on the same profile, silent SSO via `ESTSAUTHPERSISTENT`, re-capture, persist. Short timeout                                                                                                                                                                 |
| `auth-check`                             | Probe and report, see below                                                                                                                                                                                                                                                    |
| `health-check`                           | Same probes, formatted for cron with per-probe timings                                                                                                                                                                                                                         |

Retained from the existing implementation, all of which is already correct:
context-level request listening, `.mcas.ms` URL matching, treating a missing
Bearer as success rather than error on cookie-auth tenants, and cookie-derived
expiry with a conservative fallback when `FedAuth` is a session cookie.

**`auth-check` probes every dependency, not one.** This is the direct lesson from
appendix A: `teams-cli auth-check` reports `ok` on half-dead sessions because it
probes one of three backends. Here the CLI depends on three distinct surfaces, so
`auth-check` probes all three and reports per-surface status, exiting non-zero if
any fails.

| Probe  | Endpoint                              | Covers                             |
| ------ | ------------------------------------- | ---------------------------------- |
| read   | `GET /_api/web`                       | Cookie validity, site reachability |
| write  | `POST /_api/contextinfo`              | Digest minting, write capability   |
| search | `GET /_api/search/query` (rowlimit 1) | Search service entitlement         |

## 7. Modules

### 7.1 `paths.ts`

Isolated because it is the most error-prone part of SPO REST and because nearly
every NBG document has a Greek filename.

- Double apostrophes inside OData string literals (`O'Brien` becomes `O''Brien`).
  A filename containing an apostrophe otherwise terminates the literal early and
  produces a malformed query rather than a clean error.
- Percent-encode the URI while leaving the OData literal decoded. Two distinct
  encoding layers on the same string; conflating them is the classic SPO
  double-encoding bug.
- Reject `..` traversal and absolute-URL injection in caller-supplied paths.
- Split a server-relative path into parent folder and leaf name.

**Use the `*ByServerRelativePath(decodedUrl=...)` family throughout, not
`*ByServerRelativeUrl(...)`.** Both were probed and both return 200, but the
`Url` variants mishandle `#` and `%` in file names while the `Path` variants do
not. Not hypothetical: the first library listing on this tenant returns
`Βιβλιοθήκη στυλ`, `Έγγραφα`, and `Πρότυπα φόρμας`, so non-ASCII names are the
norm. The same reasoning applies to writes, where `addUsingPath(DecodedUrl=...)`
replaces `add(url=...)`.

### 7.2 `digest.ts`

- `getDigest(force?)` returns a cached digest, refetching from
  `POST /_api/contextinfo` when absent or expired.
- Expiry is `capturedAt + FormDigestTimeoutSeconds - 60s`. The margin covers
  clock skew and uploads that begin just before expiry.
- Process-local only, never persisted. It is a CSRF token with a lifetime shorter
  than most cron gaps, so writing it to disk adds attack surface for no benefit.

### 7.3 `upload.ts`

Size-routed, threshold 10 MB. Below threshold, one request:

```http
POST /_api/web/GetFolderByServerRelativePath(decodedUrl='<folder>')
     /Files/addUsingPath(DecodedUrl='<name>',overwrite=<bool>)
```

At or above threshold, chunked at 10 MB. SPO requires the target file to exist
before a chunked session starts, so: create a zero-byte file with `addUsingPath`,
then

```http
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/StartUpload(uploadId=guid'<guid>')
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/ContinueUpload(uploadId=guid'<guid>',fileOffset=<n>)
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/FinishUpload(uploadId=guid'<guid>',fileOffset=<n>)
```

`ContinueUpload` repeats until the final chunk, which goes to `FinishUpload`. The
`uploadId` is a client-generated GUID, constant for the session. On any chunk
failure the partial file is deleted so a retry starts clean rather than resuming
into an inconsistent file.

## 8. Command surface

All commands load the session and fail with the `auth_required` contract when it
is missing or expired.

| Command                                     | Endpoint                                                                                 | Notes                                                                       |
| ------------------------------------------- | ---------------------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| `ls <server-relative-path>`                 | `GET /_api/web/GetFolderByServerRelativePath(decodedUrl='<p>')?$expand=Folders,Files`    | Subfolders and files with size and modified time                            |
| `get <path\|url> [--out <file>]`            | `GET /_api/web/GetFileByServerRelativePath(decodedUrl='<p>')/$value`                     | Absolute-URL form replaces `download-sharepoint-link`                       |
| `put <local> <remote-folder> [--overwrite]` | see 7.3                                                                                  | Auto-routes small vs chunked                                                |
| `mkdir <server-relative-path>`              | `POST /_api/web/folders/addUsingPath(DecodedUrl='<p>')`                                  | Parents must exist, no implicit `-p`                                        |
| `search <query> [--rows N]`                 | `GET /_api/search/query?querytext='<q>'`                                                 | `selectproperties` limited to Title, Path, FileType, LastModifiedTime, Size |
| `libraries [site-url]`                      | `GET /_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$expand=RootFolder` | Document libraries only                                                     |

There is deliberately no `sites` command. Site discovery is
`search "contentclass:STS_Site"`, reusing machinery that must exist anyway.

Reads send `Accept: application/json;odata=nometadata`. Writes send
`X-RequestDigest` and, where `__metadata` is required,
`Content-Type: application/json;odata=verbose`.

## 9. Error handling

**SharePoint answers a stale digest with 403, not 401.** A naive status map
therefore reports every expired-digest write as a permissions failure, sending
the operator down the wrong diagnostic path. The discriminator is the string
`security validation for this page is invalid` in the response body.

| Status | Condition                          | Behaviour                                                     |
| ------ | ---------------------------------- | ------------------------------------------------------------- |
| 401    | any                                | `auth_required`, message hints at `sharepoint-cli auth-renew` |
| 403    | body matches digest-invalid marker | Refetch digest, retry once, then fail                         |
| 403    | otherwise                          | `access_denied`                                               |
| 404    | any                                | `not_found`                                                   |
| 423    | any                                | `locked`, file is checked out                                 |
| 507    | any                                | `quota_exceeded`                                              |

The retry-once-then-fail shape deliberately mirrors the 401 envelope in
`outlook-access/src/http/outlook-client.ts:435`, so all three CLIs behave the
same way under credential churn.

## 10. MCP exposure

New `plessas-marketplace/plugins/files/`, structured like
`plugins/mail/mcp-server` and `plugins/chat/mcp-server`: a thin subprocess bridge
over `sharepoint-cli`, with no independent auth or business logic.

Tools: `sharepoint_list`, `sharepoint_get`, `sharepoint_put`, `sharepoint_mkdir`,
`sharepoint_search`, `sharepoint_libraries`.

## 11. Phasing

Each phase is independently verifiable and leaves the tree working.

1. **Scaffold and auth.** Repo, tooling, pre-commit parity, `login`,
   `auth-renew`, `auth-check`, `health-check`, session store, lock. Proven by
   authenticating from a cold profile and surviving a headless renew.
2. **Read surface.** `ls`, `get`, `search`, `libraries`, plus `paths.ts`.
3. **Write surface.** `digest.ts`, `upload.ts`, `mkdir`, `put`.
4. **MCP plugin.** `plugins/files/` in `plessas-marketplace`.
5. **Removal.** Strip `--sharepoint-host`, `download-sharepoint-link`,
   `sharepoint-capture.ts`, `sharepoint-client.ts`, and `sharepoint-schema.ts`
   from `outlook-access`, along with their tests and doc entries. This spec file
   moves to `sharepoint-access` rather than being deleted, leaving a one-line
   pointer behind in the `outlook-access` CHANGELOG. Update the global
   `CLAUDE.md` re-auth line. Migrate `~/.outlook-cli/sharepoint-session.json` to
   `~/.sharepoint-cli/session.json`, or simply re-run `login`.

Phase 5 runs only after phases 1 to 3 are proven against the live tenant.

## 12. Testing

Vitest, `fetch` mocked, specs under `test_scripts/*.spec.ts`.

| Spec                        | Coverage                                                                                                                         |
| --------------------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `paths.spec.ts`             | Apostrophe doubling, Greek UTF-8 encoding, encode-layer separation, traversal rejection, parent/leaf split                       |
| `digest.spec.ts`            | Cache hit, expiry at boundary, safety margin, refetch on stale                                                                   |
| `upload.spec.ts`            | Route selection at the 10 MB boundary, chunk maths for exact multiple, remainder, and single chunk, cleanup on mid-chunk failure |
| `sharepoint-client.spec.ts` | Status-to-error mapping, including 403-digest versus 403-denied discrimination and the retry-once path                           |
| `capture.spec.ts`           | Cold-profile interactive path, warm-profile silent path, missing-Bearer-is-not-an-error, `.mcas.ms` URL matching                 |
| `auth-check.spec.ts`        | Degraded when any one of the three probes fails, ok only when all pass                                                           |

One live smoke script gated behind an environment variable so it never runs in
CI. Read-only except for a single write into a scratch folder nominated by the
operator, using a Greek-named fixture file.

## 13. Documentation obligations

Inherited from the `outlook-access` CLAUDE.md conventions, which the new repo
copies:

- `docs/design/project-design.md` as the founding design (this spec, adapted).
- `docs/design/project-functions.MD` registering every command.
- `Issues - Pending Items.md` at the repo root.
- Implementation plans under `docs/superpowers/plans/`.
- `outlook-access` docs updated in phase 5 to drop SharePoint.

## 14. Risks

| Risk                                                                          | Mitigation                                                                          |
| ----------------------------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| Cold-profile login behaves differently from the warm path currently exercised | Phase 1 is exactly this, proven before anything is built on top                     |
| Tenant policy restricts parts of classic REST not yet probed                  | Probe each endpoint against the live tenant before building its command             |
| Greek filename encoding breaks in a case unit tests miss                      | Live smoke test uses a Greek-named fixture                                          |
| Chunked upload leaves partial files on failure                                | Explicit delete of the target on any chunk error                                    |
| Digest goes stale mid-upload on a slow link                                   | 60-second safety margin plus the retry-once re-mint                                 |
| Phase 5 removal breaks an unknown consumer                                    | Three repos already searched and clean; removal is last and gated on 1 to 3 working |
| Two Playwright profiles mean two ~90-day re-logins                            | Accepted, matches `teams-access`                                                    |

---

## Appendix A: `teams-access` auth-check fix

Unrelated to SharePoint, carried here because it is a ten-line change that does
not warrant its own spec cycle, and because it motivated the `auth-check` design
in section 6.

`src/commands/auth-check.ts` probes Graph `/me` alone. The CLI depends on three
backends: Graph, chatsvc, and chatsvcagg. A live Graph token alongside a dead
chatsvc scope currently reports `status: ok`.

Fix: run the same three probes `health-check` already implements, report
per-backend status, exit non-zero when any fails. Existing response fields are
preserved and a `probes[]` array added alongside, so sentinel scripts parsing the
current output keep working. Test file is `test_scripts/auth-check.test.ts`,
noting that `teams-access` uses the `.test.ts` suffix.

Ships as an independent commit in its own repo, sequenced whenever convenient.
