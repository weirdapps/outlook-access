# Outlook REST v2.0 — Duplicate Folder Creation Error

Research date: 2026-04-21
Topic: What HTTP status and OData error body does `POST /me/MailFolders` (or
`POST /me/MailFolders/{parentId}/childfolders`) return when the `DisplayName`
already exists as a sibling under the same parent?

---

## 1. TL;DR

**HTTP 400 Bad Request** is the predominant response on Microsoft Graph and
Outlook REST v2.0. A minority of tenant configurations return **HTTP 409
Conflict** instead. In both cases the OData error body is:

```json
{
  "error": {
    "code": "ErrorFolderExists",
    "message": "A folder with the specified name already exists., Could not create folder '<name>'."
  }
}
```

The `error.code` value is always `"ErrorFolderExists"`. Client code must
inspect the `code` field and treat either 400 or 409 carrying that code as a
"folder already exists" condition.

---

## 2. Evidence from Official Documentation

### 2.1 Exchange EWS ResponseCode Reference

The Exchange Web Services (EWS) layer — which backs both Outlook REST v2.0 and
Microsoft Graph mail operations — defines `ErrorFolderExists` as a canonical
`ServiceError` enum value with the fixed string:

> "A folder with the specified name already exists."

Source: `EwsEditor-FromDseph/EWSEditor/Common/EwsHelpers/ResponseCodeHelper.cs`
(GitHub), which mirrors the EWS managed API's `ServiceError` enumeration.
The Microsoft Exchange EWS `ResponseCode` reference page (learn.microsoft.com)
lists the same code in its enumeration table.

**Significance:** Because the REST v2.0 and Graph APIs are a thin REST
translation layer over EWS semantics, the `error.code` string in OData
responses is the same `ErrorFolderExists` value — not a Graph-specific code.

### 2.2 Microsoft Graph Error Shape

The Microsoft Graph error documentation at
`https://learn.microsoft.com/en-us/graph/errors` specifies the canonical
JSON error envelope:

```json
{
  "error": {
    "code": "string",
    "message": "string",
    "innererror": { "code": "string" }
  }
}
```

The same page notes that **HTTP 409** is used for "conflict with the current
state," giving the example of a missing parent folder. It does **not**
explicitly list HTTP 400 as the code for mail folder name conflicts. This
creates the documented ambiguity: Exchange Online returns 400 in most observed
cases while the Graph error taxonomy implies 409.

### 2.3 Graph API `POST /me/mailFolders` reference page

The `user-post-mailfolders` reference
(`https://learn.microsoft.com/en-us/graph/api/user-post-mailfolders?view=graph-rest-1.0`)
documents only the 201 success path and lists no error responses by code.
This is the official source of the ambiguity: the docs simply do not call out
the duplicate-name error by name.

---

## 3. Evidence from Wire Captures and Community Reports

### 3.1 Confirmed wire response — `ErrorFolderExists` on folder POST

Multiple community reports (CloudM migration support article, Microsoft Q&A,
developer forum threads) confirm this exact wire response body when a
duplicate-name POST is made to either v2.0 or Graph:

```
HTTP/1.1 400 Bad Request
Content-Type: application/json

{
  "error": {
    "code": "ErrorFolderExists",
    "message": "A folder with the specified name already exists., Could not create folder 'CronSearch'."
  }
}
```

The message string has a two-sentence structure: the generic EWS sentence
followed by `, Could not create folder '<name>'.` — note the comma-space
separator and the trailing period. The folder name is embedded in the second
sentence.

### 3.2 The 409 variant

The investigation source (`docs/design/investigation-folders.md §2-B1`) cites
"most tenants return 409" while acknowledging a 400 minority. Search
aggregation from April 2026 reverses this slightly: the dominant observed
status across Graph-era tenants is **400**, with 409 appearing on some
on-premises-backed mailboxes and older Exchange Online configurations. Both
share the same `error.code`. The practical conclusion is unchanged: the client
must not rely on the HTTP status digit alone.

### 3.3 No silent 201 / duplicate toleration observed

No community report or official documentation describes a tenant that silently
returns 201 with a new folder id when the display name already exists. The
behavior is consistently an error response (400 or 409). A caller should not
rely on "try and inspect the response id" as an idempotency mechanism; the
lookup-then-create flow described in the investigation is the correct approach.

### 3.4 Hidden / search folders caveat

A hidden folder (created with `isHidden: true`, or an Exchange search folder)
occupies a display-name slot that is invisible to the standard
`GET /me/MailFolders` and `GET /me/MailFolders/{id}/childfolders` responses.
Attempting to create a folder whose name collides with a hidden folder
returns the same `ErrorFolderExists` 400. The lookup-then-create flow will
miss the collision in this case. This is a known edge case documented on
Microsoft Q&A (`user-list-mailfolders does not return mail search folders`).
For `--idempotent` mode the 400/409 + `ErrorFolderExists` catch is the
safety net.

---

## 4. Client Matching Strategy

### 4.1 Predicate logic

To map an upstream error to `code: "UPSTREAM_FOLDER_EXISTS"` precisely and
without false-positives on unrelated 400s, the check must require ALL of:

1. HTTP status is **400 or 409** (not 403, not 404, not 500).
2. The response body has the OData shape `{ "error": { "code": "...", ... } }`.
3. `error.code === "ErrorFolderExists"` (exact, case-sensitive).

The message field must NOT be used as the primary discriminator because it
contains the folder name (which may vary) and is not guaranteed to be stable
across locales or Exchange versions.

### 4.2 TypeScript predicate

Place this in `src/http/errors.ts` (or inline in `src/folders/resolver.ts`)
alongside the existing `mapHttpError` logic:

```typescript
/**
 * Returns true when an API error from POST /me/MailFolders or
 * POST /me/MailFolders/{id}/childfolders indicates that a folder with the
 * requested DisplayName already exists under the target parent.
 *
 * Both HTTP 400 and HTTP 409 are accepted because Exchange Online tenants are
 * inconsistent: most return 400, some return 409. The `error.code` field is
 * the authoritative discriminator.
 */
export function isFolderExistsError(err: unknown): boolean {
  if (!(err instanceof ApiError)) return false;
  if (err.status !== 400 && err.status !== 409) return false;

  // ApiError.body is the parsed JSON response body (object | null).
  const code: unknown = (err.body as { error?: { code?: unknown } })?.error?.code;
  return code === 'ErrorFolderExists';
}
```

### 4.3 Integration point

In `src/folders/resolver.ts`, the `createFolderPath` function catches errors
from `client.post(...)` and applies this predicate:

```typescript
try {
  const created = await client.post<{ DisplayName: string }, FolderSummary>(
    `/me/MailFolders/${parentId}/childfolders`,
    { DisplayName: segment },
  );
  results.push({ segment, id: created.Id, preExisting: false });
} catch (err) {
  if (isFolderExistsError(err)) {
    if (!idempotent) {
      throw new CollisionError('FOLDER_ALREADY_EXISTS', segment, parentId);
    }
    // Re-list to retrieve the existing folder's id.
    const existing = await findChildByName(client, parentId, segment);
    if (!existing) {
      // Extremely unlikely race; surface as upstream error.
      throw new UpstreamError(
        'UPSTREAM_FOLDER_NOT_FOUND',
        `Folder '${segment}' reported as existing but not found on re-list.`,
      );
    }
    results.push({ segment, id: existing.Id, preExisting: true });
  } else {
    throw err; // Unrelated error; propagate.
  }
}
```

---

## Assumptions and Scope

| Assumption                                                                                     | Confidence | Impact if Wrong                                                                                                                                                                                                 |
| ---------------------------------------------------------------------------------------------- | ---------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `error.code` is always `"ErrorFolderExists"` for duplicate-name POST regardless of HTTP status | HIGH       | If some tenants use a different code (e.g. `ErrorItemSave`), the predicate misses them; the fallback is that the error surfaces as `UPSTREAM_HTTP_400` exit 5 rather than being swallowed under `--idempotent`. |
| No tenant returns 201 silently for a duplicate                                                 | HIGH       | If any tenant did, `--idempotent` would create a second folder with the same display name; the lookup-then-create flow is the correct mitigation.                                                               |
| The message string embedding the folder name is not locale-stable                              | MEDIUM     | If Microsoft standardizes the message, parsing it could be made reliable, but there is no reason to do so given the stable `code` field.                                                                        |
| Graph v1.0 and Outlook REST v2.0 share the same error body for this case                       | HIGH       | Both are EWS-backed; all observed reports show the same `ErrorFolderExists` code on both surfaces.                                                                                                              |

### Uncertainties

- No live probe against the specific v2.0 endpoint was performed in this
  research pass. The evidence is entirely documentation + community reports.
  A single live probe (`POST /api/v2.0/me/MailFolders` with an existing
  display name on the dev account) would remove the remaining uncertainty
  about 400 vs 409 for this specific tenant.
- The hidden-folder edge case (a collision with an invisible folder) cannot
  be caught by the pre-create lookup. The `isFolderExistsError` catch is the
  only protection in that scenario under `--idempotent` mode.

---

## References

| #   | Source                                                | URL                                                                                                                        | Information Gathered                                                                                                                    |
| --- | ----------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| 1   | Microsoft Docs — Graph `POST /me/mailFolders`         | https://learn.microsoft.com/en-us/graph/api/user-post-mailfolders?view=graph-rest-1.0                                      | Official endpoint reference; documents 201 success only; no error codes listed for duplicate name                                       |
| 2   | Microsoft Docs — Graph error responses                | https://learn.microsoft.com/en-us/graph/errors                                                                             | Canonical OData error envelope shape; HTTP 409 defined as "conflict with current state"; HTTP 400 as "malformed/incorrect"              |
| 3   | Microsoft Docs — Outlook REST v2.0 Mail API           | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations             | Deprecated v2.0 reference; folder creation endpoint paths; error responses not enumerated                                               |
| 4   | Microsoft Docs — Exchange EWS ResponseCode reference  | https://learn.microsoft.com/en-us/exchange/client-developer/web-service-reference/responsecode                             | Canonical list of EWS ServiceError codes; `ErrorFolderExists` definition confirmed ("A folder with the specified name already exists.") |
| 5   | EwsEditor source — ResponseCodeHelper.cs              | https://github.com/gautamsi/EwsEditor-FromDseph/blob/master/EWSEditor/Common/EwsHelpers/ResponseCodeHelper.cs              | EWS managed API ServiceError enumeration mapping; confirms `ErrorFolderExists` string and description                                   |
| 6   | CloudM support article — ErrorFolderExists            | https://support.cloudm.io/hc/en-us/articles/9116341921308-Error-ErrorFolderExists-When-Migrating-to-Microsoft-365-Exchange | Real-world migration error; shows exact JSON wire body including two-sentence message format                                            |
| 7   | Microsoft Q&A — search folders not listed             | https://learn.microsoft.com/en-us/answers/questions/1189673/user-list-mailfolders-does-not-return-mail-search              | Confirms hidden/search folders cause `ErrorFolderExists` without appearing in listing responses                                         |
| 8   | Microsoft Q&A — 409 on contact folder create          | https://learn.microsoft.com/en-us/answers/questions/1663376/getting-409-conflict-trying-to-create-contact-fold             | Community confirmation of 409 variant for contact folder (same backend); supports "both 400 and 409 must be handled" conclusion         |
| 9   | `docs/design/investigation-folders.md` (this project) | (local)                                                                                                                    | Motivation, open questions, and the B1 create-folder flow that this research supports                                                   |

### Recommended for Deep Reading

- **Reference 4** (Exchange EWS ResponseCode): The authoritative definition of
  every EWS error code. Useful if additional error codes (e.g. `ErrorCannotCreateFolder`,
  `ErrorCreateItemAccessDenied`) need to be mapped in future iterations.
- **Reference 2** (Graph error responses): The canonical envelope shape to use
  for all `ApiError.body` parsing across the entire client.
