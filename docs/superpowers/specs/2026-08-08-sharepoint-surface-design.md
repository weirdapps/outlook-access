# SharePoint Command Surface, and teams-access auth-check Fix

**Date:** 2026-08-08
**Repos:** `outlook-access` (primary), `teams-access` (one contained fix), `plessas-marketplace` (new `files` plugin)
**Status:** Approved design, pending implementation plan

---

## 1. Problem

`outlook-access` captures a working SharePoint session (`login --sharepoint-host`,
refreshed silently by `auth-renew --sharepoint-host`) but exposes almost nothing
on top of it. `SharepointClient` has exactly one method, `getBinary(absoluteUrl)`,
surfaced as one command, `download-sharepoint-link`. There is no way to browse a
document library, list a folder, download by path, search, create a folder, or
upload.

The auth layer is not the gap. The capture path already listens at
`context.on('request')`, which catches Service-Worker-dispatched and MCAS-proxied
requests that a page-level or in-page hook misses. Session persistence, the
`.browser.lock` concurrency guard, and headless silent renewal all work. The gap
is purely API surface.

Separately, `teams-cli auth-check` reports `status: ok` on sessions that are
partially dead, because it probes only one of the three backends the CLI depends
on. This has already cost debugging time and is recorded as a known quirk.

## 2. Established constraints

These were measured against the live tenant SharePoint session on
2026-08-08, not assumed. They drive every subsequent decision.

| Probe                       | Result                                    | Consequence                              |
| --------------------------- | ----------------------------------------- | ---------------------------------------- |
| Session contents            | `FedAuth` + `rtFa` cookies, **no Bearer** | Cookie auth is the only credential       |
| `GET /_api/web`             | 200                                       | Classic SPO REST is available            |
| `GET /_api/v2.0/me/drive`   | 403 `accessDenied`                        | Graph-shaped `driveItem` API is blocked  |
| `GET /_api/v2.0/sites/root` | 403 `accessDenied`                        | Same                                     |
| `POST /_api/contextinfo`    | 200, `FormDigestTimeoutSeconds: 1800`     | Writes are feasible, digest lives 30 min |
| `GET /_api/search/query`    | 200                                       | Search and site discovery are available  |

Two conclusions follow and are not revisitable without new evidence:

1. **Microsoft Graph is unusable for SharePoint here.** `graph.microsoft.com`
   requires a Graph-audience Bearer. The tenant issues none for SharePoint, and
   cookies do not authorise Graph. The tenant additionally blocks the
   SharePoint-hosted `/_api/v2.0` mirror of the same object model.
2. **Writes need CSRF handling.** Cookie-authenticated POSTs to SPO REST are
   rejected without an `X-RequestDigest` header. Bearer callers are exempt from
   this, which is why the requirement is easy to miss when reading Graph-oriented
   documentation.

## 3. Non-goals

- No change to capture, session schema, renewal, or locking. They work.
- No long-lived local proxy. A resident browser plus an unauthenticated
  `127.0.0.1` port fronting the whole mailbox is worse than the current
  persist-and-renew model, which is stateless between invocations, survives
  reboot, and suits unattended VPS cron.
- No replacement of `download-sharepoint-link`. It resolves share URLs through a
  different code path and keeps working unchanged.
- No OneDrive-for-Business personal-site support in this pass.

## 4. Architecture

```text
outlook-access/src/
  auth/sharepoint-capture.ts      unchanged
  session/sharepoint-schema.ts    unchanged
  http/sharepoint-client.ts       EXPANDED: getBinary -> full REST client
  sharepoint/                     NEW
    paths.ts                      server-relative path handling
    digest.ts                     X-RequestDigest fetch + cache
    upload.ts                     size-routed upload strategy
  commands/
    sp-ls.ts  sp-get.ts  sp-put.ts  sp-mkdir.ts  sp-search.ts  sp-libraries.ts
```

Each new module has one job, no shared mutable state beyond the digest cache,
and can be tested with `fetch` mocked.

### 4.1 `paths.ts`

Server-relative path handling, isolated because it is the single most
error-prone part of SPO REST and because nearly every NBG document has a Greek
filename.

Responsibilities:

- Double apostrophes inside OData string literals (`O'Brien` becomes `O''Brien`).
  A Greek or English filename containing an apostrophe otherwise terminates the
  literal early and produces a malformed query rather than a clean error.
- Percent-encode the URI while leaving the OData literal decoded. These are two
  distinct encoding layers applied to the same string, and conflating them is the
  classic SPO double-encoding bug.
- Reject `..` traversal and absolute-URL injection in caller-supplied paths.
- Split a server-relative path into parent folder and leaf name for upload and
  mkdir.

**Use the `*ByServerRelativePath(decodedUrl=...)` family throughout, not
`*ByServerRelativeUrl(...)`.** Both were probed against the tenant on 2026-08-08
and both return 200, but the `Url` variants mishandle `#` and `%` in file names
while the `Path` variants do not. This is not a hypothetical concern: the first
library listing on this tenant returns `Βιβλιοθήκη στυλ`, `Έγγραφα`, and
`Πρότυπα φόρμας`, so non-ASCII names are the norm rather than the exception.
The same reasoning applies to writes, where `addUsingPath(DecodedUrl=...)`
replaces the older `add(url=...)`.

### 4.2 `digest.ts`

- `getDigest(force?: boolean)` returns a cached digest, refetching from
  `POST /_api/contextinfo` when absent or expired.
- Expiry is `capturedAt + FormDigestTimeoutSeconds - 60s`. The 60-second margin
  covers clock skew and slow uploads that begin just before expiry.
- Cache is process-local only. It is never persisted: it is a CSRF token whose
  lifetime is shorter than most cron gaps, so writing it to disk would add
  attack surface for no benefit.

### 4.3 `upload.ts`

Size-routed, threshold 10 MB.

Below threshold, single request:

```http
POST /_api/web/GetFolderByServerRelativePath(decodedUrl='<folder>')
     /Files/addUsingPath(DecodedUrl='<name>',overwrite=<bool>)
```

At or above threshold, chunked at 10 MB. SPO requires the target file to exist
before a chunked session starts, so the sequence is: create a zero-byte file with
`addUsingPath`, then

```http
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/StartUpload(uploadId=guid'<guid>')
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/ContinueUpload(uploadId=guid'<guid>',fileOffset=<n>)
POST .../GetFileByServerRelativePath(decodedUrl='<path>')/FinishUpload(uploadId=guid'<guid>',fileOffset=<n>)
```

`ContinueUpload` repeats until the final chunk, which goes to `FinishUpload`. The
`uploadId` is a client-generated GUID, constant for the session. On any chunk
failure the partial file is deleted so a retry starts clean rather than resuming
into an inconsistent file.

## 5. Command surface

All commands read the existing SharePoint session file and fail with the current
`auth_required` contract when it is missing or expired.

| Command                                        | Endpoint                                                                                 | Notes                                                                       |
| ---------------------------------------------- | ---------------------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| `sp-ls <server-relative-path>`                 | `GET /_api/web/GetFolderByServerRelativePath(decodedUrl='<p>')?$expand=Folders,Files`    | Lists subfolders and files with size and modified time                      |
| `sp-get <path\|url> [--out <file>]`            | `GET /_api/web/GetFileByServerRelativePath(decodedUrl='<p>')/$value`                     | Falls back to the existing absolute-URL path when given a URL               |
| `sp-put <local> <remote-folder> [--overwrite]` | see 4.3                                                                                  | Auto-routes small vs chunked                                                |
| `sp-mkdir <server-relative-path>`              | `POST /_api/web/folders/addUsingPath(DecodedUrl='<p>')`                                  | Parents must exist, no implicit `-p`                                        |
| `sp-search <query> [--rows N]`                 | `GET /_api/search/query?querytext='<q>'`                                                 | `selectproperties` limited to Title, Path, FileType, LastModifiedTime, Size |
| `sp-libraries [site-url]`                      | `GET /_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$expand=RootFolder` | Document libraries only                                                     |

There is deliberately no `sp-sites`. Site discovery is
`sp-search "contentclass:STS_Site"`, which reuses machinery that already has to
exist rather than adding a seventh command for a rare operation.

Reads send `Accept: application/json;odata=nometadata`. Writes send
`X-RequestDigest` and, where `__metadata` is required, `Content-Type:
application/json;odata=verbose`.

## 6. Error handling

Extends the existing `SharepointHttpError` rather than introducing a new type.

**The important subtlety: SharePoint answers a stale digest with 403, not 401.**
A naive status-to-error map therefore reports every expired-digest write as a
permissions failure, which sends the operator down the wrong diagnostic path. The
discriminator is the string `security validation for this page is invalid` in
the response body.

| Status | Condition                          | Behaviour                                                        |
| ------ | ---------------------------------- | ---------------------------------------------------------------- |
| 401    | any                                | `auth_required`, message hints at `auth-renew --sharepoint-host` |
| 403    | body matches digest-invalid marker | Refetch digest, retry once, then fail                            |
| 403    | otherwise                          | `access_denied`                                                  |
| 404    | any                                | `not_found`                                                      |
| 423    | any                                | `locked`, file is checked out                                    |
| 507    | any                                | `quota_exceeded`                                                 |

The retry-once-then-fail shape deliberately mirrors the 401 envelope already in
`src/http/outlook-client.ts:435`, so both clients behave the same way under
credential churn.

## 7. teams-access `auth-check` fix

Independent commit, separate repo, included here because a ten-line change does
not warrant its own spec cycle.

`src/commands/auth-check.ts` currently probes Graph `/me` alone. The CLI depends
on three backends: Graph, chatsvc, and chatsvcagg. A live Graph token alongside a
dead chatsvc scope currently reports `status: ok`.

Fix: run the same three probes `health-check` already implements, report
per-backend status, and exit non-zero when any probe fails. Existing response
fields are preserved and a `probes[]` array is added alongside them, so sentinel
scripts that parse the current output keep working without modification.

## 8. MCP exposure

New `plessas-marketplace/plugins/files/`, structured like the existing
`plugins/mail/mcp-server`: a thin subprocess bridge over the CLI, no independent
auth or business logic.

Tools: `sharepoint_list`, `sharepoint_get`, `sharepoint_put`, `sharepoint_mkdir`,
`sharepoint_search`, `sharepoint_libraries`.

A separate plugin rather than an extension of `mail` keeps the mail MCP focused
on mail and lets document tooling be enabled independently.

## 9. Testing

Vitest, `fetch` mocked. Note the two repos use different file suffixes, both
enforced by their own `vitest.config.ts`: `outlook-access` includes
`test_scripts/**/*.spec.ts`, `teams-access` includes `test_scripts/**/*.test.ts`.

| Spec                        | Repo           | Coverage                                                                                                                         |
| --------------------------- | -------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `sharepoint-paths.spec.ts`  | outlook-access | Apostrophe doubling, Greek UTF-8 encoding, encode-layer separation, traversal rejection, parent/leaf split                       |
| `sharepoint-digest.spec.ts` | outlook-access | Cache hit, expiry at boundary, safety margin, refetch on stale                                                                   |
| `sharepoint-upload.spec.ts` | outlook-access | Route selection at the 10 MB boundary, chunk maths for exact multiple, remainder, and single chunk, cleanup on mid-chunk failure |
| `sharepoint-client.spec.ts` | outlook-access | Status-to-error mapping, including 403-digest versus 403-denied discrimination and the retry-once path                           |
| `auth-check.test.ts`        | teams-access   | Reports degraded when any one backend fails, ok only when all three pass                                                         |

One live smoke script, gated behind an environment variable so it never runs in
CI. Read-only except for a single write into a scratch folder nominated by the
operator.

## 10. Documentation obligations

Required by the repo's own `CLAUDE.md`:

- Update `docs/design/project-design.md` with the SharePoint surface.
- Register the new commands in `docs/design/project-functions.MD`.
- Log any defect found during implementation in `Issues - Pending Items.md`.
- The implementation plan goes to `docs/superpowers/plans/`, following the
  precedent of `2026-04-22-send-mail-b1-core.md`.

## 11. Risks

| Risk                                                              | Mitigation                                                                         |
| ----------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| Tenant policy also restricts parts of classic REST not yet probed | Probe each endpoint against the live tenant before building its command, not after |
| Greek filename encoding breaks in a case the unit tests miss      | Live smoke test uses a Greek-named fixture file                                    |
| Chunked upload leaves partial files on failure                    | Explicit delete of the target on any chunk error                                   |
| Digest cache goes stale mid-upload on a slow link                 | 60-second safety margin, plus the retry-once path re-mints on 403                  |
| MCAS changes the proxying behaviour                               | Capture layer already handles `.mcas.ms` rewriting and is unchanged by this work   |
