# Outlook REST v2 Attachments — Implementation Reference

Research date: 2026-04-21  
Author: Technical Research Agent  
Scope: `outlook.office.com/api/v2.0` attachment endpoints — shapes, download semantics, filename safety, error mapping.

---

## Overview

The Outlook REST v2.0 API exposes attachments as a polymorphic collection under each message. The three concrete subtypes — `FileAttachment`, `ItemAttachment`, and `ReferenceAttachment` — share a common base resource but differ substantially in shape and in what the implementer should do when they appear. This document codifies everything needed to implement the `download-attachments` command without ambiguity.

**Important deprecation note:** Microsoft announced the formal deprecation of the v2.0 Outlook REST endpoint on 2020-11-17. The documented decommission date was March 2024. Despite this, the endpoint continues to respond for sessions authenticated through `outlook.office.com` web client tokens (the mechanism this CLI relies on). The API shape documented here matches what the live endpoint returns today and is also structurally identical to the Microsoft Graph v1.0 equivalents, so a future migration path to Graph is a drop-in rename of the base URL and field-name casing only.

---

## 1. API Paths

### 1.1 List all attachments on a message

```
GET https://outlook.office.com/api/v2.0/me/messages/{messageId}/attachments
Authorization: Bearer {token}
Accept: application/json
```

Returns an OData collection (`value` array) containing every attachment on the message. The `@odata.type` discriminator is present on every item. `ContentBytes` **is** returned on `FileAttachment` items in the list response, but it may be `null` or absent when the attachment size approaches or exceeds the REST request ceiling (~3-4 MB in base64 terms). Always proceed to the detail call for any `FileAttachment`.

### 1.2 Get a single attachment (detail)

```
GET https://outlook.office.com/api/v2.0/me/messages/{messageId}/attachments/{attachmentId}
Authorization: Bearer {token}
Accept: application/json
```

This is the reliable path for reading `ContentBytes`. It returns the full shape for whichever subtype matched. The detail call is always required before decoding because:
- The list endpoint may omit `ContentBytes` for larger files.
- The list endpoint's `ContentBytes` field is the one most likely to be `null` silently for attachments larger than ~3 MB (binary), which is approximately 4 MB base64-encoded.

### 1.3 Get raw binary content via `$value`

```
GET https://outlook.office.com/api/v2.0/me/messages/{messageId}/attachments/{attachmentId}/$value
Authorization: Bearer {token}
```

On **Microsoft Graph v1.0** this is officially documented and returns raw bytes with the attachment's original `Content-Type`. On `outlook.office.com/api/v2.0`, `$value` is **not officially documented** for this endpoint and should be treated as untested. The v2.0 API did not surface `/$value` in its published specification. **Do not rely on `/$value` for this CLI.** Use the JSON detail endpoint and decode `ContentBytes` from base64.

> Note for `ItemAttachment`: on Graph, `/$value` returns MIME content (RFC 822 for messages, vCard for contacts, iCal for events). This is the only way to get raw bytes out of an item attachment. On v2.0 the behaviour is unspecified; the recommended action for `ItemAttachment` remains: skip it.

### 1.4 `$expand` on `ItemAttachment`

```
GET https://outlook.office.com/api/v2.0/me/messages/{messageId}/attachments/{attachmentId}?$expand=Microsoft.OutlookServices.ItemAttachment/Item
```

This returns the nested `Item` object (Subject, Body, Sender, etc.) but does not provide downloadable bytes. It is only useful if you later decide to surface item attachments as `.eml` — out of scope for this iteration.

---

## 2. `@odata.type` Discriminator Values

The v2.0 Outlook REST API uses the following exact string values as the `@odata.type` discriminator. These differ from Microsoft Graph which uses `microsoft.graph.*` namespace:

| Attachment Kind | v2.0 `@odata.type` | Graph `@odata.type` |
|---|---|---|
| File | `#Microsoft.OutlookServices.FileAttachment` | `#microsoft.graph.fileAttachment` |
| Item | `#Microsoft.OutlookServices.ItemAttachment` | `#microsoft.graph.itemAttachment` |
| Reference | `#Microsoft.OutlookServices.ReferenceAttachment` | `#microsoft.graph.referenceAttachment` |

The leading `#` is part of the literal string value. Code that compares `@odata.type` must include the `#`.

---

## 3. JSON Shapes Per Attachment Kind

### 3.1 FileAttachment

**List endpoint response item:**

```json
{
  "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
  "@odata.id": "https://outlook.office.com/api/v2.0/Users('ddfcd489-628b-40d7-b48b-57002df800e5@1717622f-1d94-4d0c-9d74-709fad664b77')/Messages('AAMkAGI2THVSAAA=')/Attachments('AAMkAGI2THVSAAABEgAQAMkp=')",
  "Id": "AAMkAGI2THVSAAABEgAQAMkp=",
  "LastModifiedDateTime": "2024-03-12T09:31:00Z",
  "Name": "Q1-Report.pdf",
  "ContentType": "application/pdf",
  "Size": 147210,
  "IsInline": false,
  "ContentId": null,
  "ContentLocation": null,
  "ContentBytes": "JVBERi0xLjQKJeLjz9MKMSAwIG9iag..."
}
```

**Detail endpoint response** (same fields, `ContentBytes` guaranteed populated):

```json
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Me/Messages('AAMkAGI2THVSAAA=')/Attachments/$entity",
  "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
  "@odata.id": "https://outlook.office.com/api/v2.0/Users('ddfcd489...')/Messages('AAMkAGI2THVSAAA=')/Attachments('AAMkAGI2THVSAAABEgAQAMkp=')",
  "Id": "AAMkAGI2THVSAAABEgAQAMkp=",
  "LastModifiedDateTime": "2024-03-12T09:31:00Z",
  "Name": "Q1-Report.pdf",
  "ContentType": "application/pdf",
  "Size": 147210,
  "IsInline": false,
  "ContentId": null,
  "ContentLocation": null,
  "ContentBytes": "JVBERi0xLjQKJeLjz9MKMSAwIG9iag..."
}
```

**Field-by-field explanation:**

| Field | Type | Notes |
|---|---|---|
| `@odata.type` | string | Discriminator. Always `#Microsoft.OutlookServices.FileAttachment`. |
| `@odata.id` | string | Fully qualified self-link. Usable as a fetch URL. |
| `Id` | string | Opaque attachment identifier. Stable within the session; use for per-attachment GET. |
| `LastModifiedDateTime` | ISO 8601 string | UTC. Suitable for display; not needed for download. |
| `Name` | string | Display name — **not** a safe filesystem path. See Section 5. |
| `ContentType` | string | MIME type. May be `application/octet-stream` even if the file is clearly a PDF. Do not trust for extension inference. |
| `Size` | number (Int32) | Declared size in bytes of the raw (binary) content. The base64 encoding in `ContentBytes` is ~33% larger. |
| `IsInline` | boolean | `true` for inline images embedded in HTML bodies (e.g. logo images in email signatures). Default CLI behavior: skip. |
| `ContentId` | string or null | CID used in MIME parts to reference inline attachments from HTML body (`cid:` scheme). Usually non-null when `IsInline` is true. |
| `ContentLocation` | string or null | Rarely populated. A URL hint from the original MIME message. |
| `ContentBytes` | string (base64) | The raw file contents, base64-encoded. May be `null` on the list endpoint for larger files. Always populated on the detail endpoint. |

**Inline attachment identification:** An attachment is inline when `IsInline === true`. Additionally, when `ContentId` is non-null the attachment is referenced from the message HTML body via a `cid:` URI. Both conditions consistently co-occur but checking `IsInline` alone is sufficient for the CLI's skip logic.

### 3.2 ItemAttachment

An `ItemAttachment` is another Outlook item (message, calendar event, or contact) attached to the message. It carries no `ContentBytes` field.

**List/detail endpoint response:**

```json
{
  "@odata.type": "#Microsoft.OutlookServices.ItemAttachment",
  "@odata.id": "https://outlook.office.com/api/v2.0/Users('ddfcd489...')/Messages('AAMkAGI2THVSAAA=')/Attachments('AAMkAGE1Mbs88AADUv0uFAAABEgAQAL53=')",
  "Id": "AAMkAGE1Mbs88AADUv0uFAAABEgAQAL53=",
  "LastModifiedDateTime": "2024-03-10T14:05:55Z",
  "Name": "RE: Project Alpha — Meeting Notes",
  "ContentType": null,
  "Size": 78927,
  "IsInline": false,
  "Item": null
}
```

When `$expand=Microsoft.OutlookServices.ItemAttachment/Item` is added:

```json
{
  "@odata.type": "#Microsoft.OutlookServices.ItemAttachment",
  "Id": "AAMkAGE1Mbs88AADUv0uFAAABEgAQAL53=",
  "Name": "RE: Project Alpha — Meeting Notes",
  "ContentType": null,
  "Size": 78927,
  "IsInline": false,
  "Item": {
    "@odata.type": "#Microsoft.OutlookServices.Message",
    "Id": "",
    "Subject": "RE: Project Alpha — Meeting Notes",
    "Sender": {
      "EmailAddress": {
        "Name": "Alice Smith",
        "Address": "alice@contoso.com"
      }
    },
    "ReceivedDateTime": "2024-03-09T11:22:00Z"
  }
}
```

**Key notes for the implementer:**
- `ContentBytes` is never present on `ItemAttachment`. There is no field to decode.
- `Item` is `null` unless `$expand` is used.
- The attached item may itself be a `Message`, `Event`, or `Contact` — the `@odata.type` inside the `Item` object identifies which.
- **Implementation decision:** Skip `ItemAttachment` entries. Record in the `skipped[]` output with `reason: "item-attachment"` and include `Name` and `Id` so the user is informed.

### 3.3 ReferenceAttachment

A `ReferenceAttachment` is a hyperlink to a cloud-stored file (OneDrive, OneDrive for Business, SharePoint, or Dropbox). There are no bytes to download; the API provides a URL to the external resource.

**List/detail endpoint response:**

```json
{
  "@odata.type": "#Microsoft.OutlookServices.ReferenceAttachment",
  "@odata.id": "https://outlook.office.com/api/v2.0/Users('ddfcd489...')/Messages('AAMkAGI2THVSAAA=')/Attachments('AAMkAGI2THVSAAABEgAQAPSg=')",
  "Id": "AAMkAGI2THVSAAABEgAQAPSg=",
  "LastModifiedDateTime": "2024-03-12T06:04:38Z",
  "Name": "Q1 Budget Model",
  "ContentType": null,
  "Size": 382,
  "IsInline": false,
  "SourceUrl": "https://contoso-my.sharepoint.com/personal/alice_contoso_com/Documents/Budget/Q1-2024.xlsx",
  "ProviderType": "OneDriveBusiness",
  "ThumbnailUrl": null,
  "PreviewUrl": null,
  "Permission": "Edit",
  "IsFolder": false
}
```

**Field-by-field explanation:**

| Field | Type | Notes |
|---|---|---|
| `SourceUrl` | string | The full URL to the file or folder in the cloud storage provider. Include in the `skipped[]` record so the user can access it manually. |
| `ProviderType` | string | One of: `oneDriveBusiness`, `oneDriveConsumer`, `dropbox`, `box`, `google`, `other`. Informational only. |
| `Permission` | string | `"Edit"` or `"View"`. Describes the access level the attachment link conveys. |
| `IsFolder` | boolean | `true` if the link points to a folder rather than a file. |
| `ThumbnailUrl` | string or null | Often null. |
| `PreviewUrl` | string or null | Often null. |

**Implementation decision:** Skip `ReferenceAttachment` entries. Record in the `skipped[]` output with `reason: "reference-attachment"` and include `Name`, `Id`, and `SourceUrl`.

---

## 4. Size Limits and Large Attachment Handling

### 4.1 The 4 MB ceiling on `ContentBytes`

The REST JSON response body has an effective ceiling of 4 MB per request. `ContentBytes` is base64-encoded, which inflates binary content by approximately 33%. As a result:

- A binary file of ~3 MB encodes to ~4 MB base64 — approaching the ceiling.
- A binary file of >3 MB may cause the API to return `ContentBytes: null` on the list endpoint, and may cause a 413 or truncated response on the detail endpoint.

**Safe range:** For files with `Size <= 3 * 1024 * 1024` (3 MB), the detail endpoint's `ContentBytes` is reliable. For files with `Size > 3 MB`, the behavior is server-dependent.

### 4.2 Behavior when `Size` exceeds ~3 MB

The v2.0 Outlook REST API does not provide a chunked download or range-request mechanism for reading attachment bytes on the GET endpoint. (The upload-session API exists for `PUT`, not `GET`.) Options when `Size > 3 MB`:

1. **Attempt the detail GET anyway.** The server may still return `ContentBytes` for large attachments depending on the Exchange tenant configuration. If `ContentBytes` is non-null, decode and write normally. If `ContentBytes` is null, fall through to option 2.

2. **Record as oversized-skipped.** If `ContentBytes` is null after the detail GET and `Size > LARGE_ATTACHMENT_THRESHOLD_BYTES`, add a `skipped[]` entry with `reason: "content-bytes-null"` and `size` so the user can see what was not downloaded.

3. **Future: `$value` on Graph.** If this CLI ever migrates to Microsoft Graph, `GET /me/messages/{id}/attachments/{attId}/$value` returns raw binary without the 4 MB base64 ceiling and works correctly for large files.

### 4.3 Recommended constants

```typescript
// Attachments larger than this will have ContentBytes=null on the list endpoint.
// Always use the detail endpoint regardless; this threshold gates the "too large" warning.
const LARGE_ATTACHMENT_BYTES = 3 * 1024 * 1024; // 3 MB binary
```

---

## 5. Filename Sanitization

Attachment `Name` fields come from the sender and must be treated as untrusted input. The following threat vectors must be handled:

**Path traversal:** Names like `../../etc/passwd` or `..\Windows\system32\config` use directory separator characters to escape the intended output directory.

**Windows reserved device names:** `CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9` are special on Windows and cannot be used as filenames regardless of extension (e.g. `NUL.pdf` is still dangerous on Windows).

**Empty or whitespace-only names:** If `Name` is empty, `null`, or only whitespace, a fallback synthetic name must be generated.

**Duplicate names:** When a message has two attachments with the same `Name`, writing the second would overwrite the first. A deduplication suffix must be applied.

**Control characters and null bytes:** Characters 0x00–0x1F (including `\0`, `\n`, `\r`, `\t`) are illegal in filenames on most OSes or cause silent truncation.

**Unicode normalization:** Some Unicode characters are visually identical to ASCII but map to different code points; normalize to NFC first to prevent encoding-based evasion.

**Excessively long names:** Most filesystems cap filenames at 255 bytes (not characters). Truncate safely to leave room for deduplication suffixes.

### 5.1 Complete TypeScript sanitization routine

```typescript
import path from 'node:path';

/** Windows reserved device names (case-insensitive, any extension). */
const WINDOWS_RESERVED = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)/i;

/** Characters illegal in filenames on Windows or POSIX: / \ : * ? " < > | and control chars. */
const ILLEGAL_CHARS = /[/\\:*?"<>|\x00-\x1F]/g;

/** Maximum byte-length of a filename component (reserve 12 bytes for suffix like " (99).ext"). */
const MAX_FILENAME_BYTES = 243;

/**
 * Sanitize an attachment Name field into a safe filesystem filename.
 *
 * @param raw       - The Name field value from the API (may be null/undefined).
 * @param fallback  - Used when raw is empty after sanitization (e.g. "attachment-1").
 * @returns         A filename safe for use on both POSIX and Windows.
 */
export function sanitizeAttachmentName(raw: string | null | undefined, fallback: string): string {
  // 1. Coerce null/undefined to empty string.
  let name = (raw ?? '').trim();

  // 2. Unicode NFC normalization — prevent lookalike bypass.
  name = name.normalize('NFC');

  // 3. Strip illegal characters (path separators, control chars, Windows-forbidden chars).
  name = name.replace(ILLEGAL_CHARS, '_');

  // 4. Strip leading dots (hidden files on POSIX) and trailing dots/spaces (illegal on Windows).
  name = name.replace(/^\.+/, '').replace(/[\s.]+$/, '');

  // 5. Use fallback if the result is now empty.
  if (name.length === 0) {
    name = fallback;
  }

  // 6. Reject Windows reserved device names (replace the base name, keep the extension).
  if (WINDOWS_RESERVED.test(name)) {
    const ext = path.extname(name);
    name = `_reserved_${name.slice(0, name.length - ext.length)}${ext}`;
  }

  // 7. Enforce maximum byte length (UTF-8 bytes, not character count).
  const encoder = new TextEncoder();
  let encoded = encoder.encode(name);
  if (encoded.byteLength > MAX_FILENAME_BYTES) {
    const ext = path.extname(name);
    const extBytes = encoder.encode(ext).byteLength;
    const baseName = name.slice(0, name.length - ext.length);
    // Truncate the base name to fit.
    let truncated = baseName;
    while (encoder.encode(truncated + ext).byteLength > MAX_FILENAME_BYTES) {
      truncated = truncated.slice(0, -1);
    }
    name = truncated + ext;
  }

  return name;
}

/**
 * Given a target directory and a desired filename, return a path that does not
 * already exist by appending " (N)" before the extension when necessary.
 *
 * @param dir      - The absolute output directory (must already exist).
 * @param filename - The sanitized (but potentially duplicate) filename.
 * @returns        An absolute path guaranteed not to exist at the time of the call.
 */
export function deduplicateFilename(dir: string, filename: string): string {
  const ext = path.extname(filename);
  const base = filename.slice(0, filename.length - ext.length);

  let candidate = path.join(dir, filename);
  // Verify the candidate does not escape the intended directory (defense in depth).
  const resolved = path.resolve(candidate);
  const resolvedDir = path.resolve(dir);
  if (!resolved.startsWith(resolvedDir + path.sep) && resolved !== resolvedDir) {
    throw new Error(`Path traversal detected: "${filename}" resolves outside output directory`);
  }

  let counter = 1;
  const { existsSync } = await import('node:fs');
  while (existsSync(candidate)) {
    candidate = path.join(dir, `${base} (${counter})${ext}`);
    counter++;
    if (counter > 999) {
      throw new Error(`Cannot find a unique filename for "${filename}" after 999 attempts`);
    }
  }

  return candidate;
}
```

> Note: `deduplicateFilename` uses `existsSync` for simplicity. In the CLI's actual write path this check + write must be wrapped in the `open(0o600) + rename` atomic write pattern described in `investigation-outlook-cli.md §4.8` to avoid TOCTOU races.

---

## 6. Download-All-Attachments Pseudocode

The following pseudocode describes the complete `download-attachments` command loop. It is written to match the error taxonomy and exit codes in `investigation-outlook-cli.md §4.9`.

```typescript
async function downloadAttachments(
  messageId: string,
  outputDir: string,
  opts: { includeInline: boolean; overwrite: boolean }
): Promise<DownloadResult> {
  // ── Step 1: List attachments ────────────────────────────────────────────────
  const listUrl = `https://outlook.office.com/api/v2.0/me/messages/${messageId}/attachments`;
  const listResp = await apiGet(listUrl);           // throws UpstreamError on 4xx/5xx
  const attachments: AttachmentEnvelope[] = listResp.value;

  const downloaded: DownloadedRecord[] = [];
  const skipped: SkippedRecord[] = [];

  // ── Step 2: Process each attachment ────────────────────────────────────────
  for (const att of attachments) {
    const odataType: string = att['@odata.type'] ?? '';

    // ── 2a: ReferenceAttachment — no bytes, always skip ──────────────────────
    if (odataType === '#Microsoft.OutlookServices.ReferenceAttachment') {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'reference-attachment',
        sourceUrl: att.SourceUrl ?? null,
      });
      continue;
    }

    // ── 2b: ItemAttachment — embedded Outlook item, skip in this iteration ───
    if (odataType === '#Microsoft.OutlookServices.ItemAttachment') {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'item-attachment',
      });
      continue;
    }

    // ── 2c: Unknown type — skip defensively ──────────────────────────────────
    if (odataType !== '#Microsoft.OutlookServices.FileAttachment') {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'unknown-attachment-type',
        odataType,
      });
      continue;
    }

    // ── 2d: FileAttachment ───────────────────────────────────────────────────

    // Inline check (before fetching detail to save an API call).
    if (att.IsInline === true && !opts.includeInline) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'inline',
      });
      continue;
    }

    // ── Step 3: Fetch detail to get ContentBytes reliably ────────────────────
    const detailUrl = `https://outlook.office.com/api/v2.0/me/messages/${messageId}/attachments/${att.Id}`;
    let detail: FileAttachmentDetail;
    try {
      detail = await apiGet(detailUrl);
    } catch (err) {
      if (err instanceof UpstreamError && err.httpStatus === 404) {
        // Attachment was deleted between list and detail GET.
        skipped.push({ id: att.Id, name: att.Name, reason: 'not-found' });
        continue;
      }
      if (err instanceof UpstreamError && err.httpStatus === 403) {
        // Item-level access denied (e.g. IRM-protected attachment).
        skipped.push({ id: att.Id, name: att.Name, reason: 'access-denied' });
        continue;
      }
      throw err; // All other errors propagate.
    }

    // ── Step 4: Handle null ContentBytes (large attachment) ──────────────────
    if (detail.ContentBytes == null) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'content-bytes-null',
        size: att.Size,
        hint: 'Attachment may exceed the 3 MB REST limit. Download manually.',
      });
      continue;
    }

    // ── Step 5: Decode base64 ────────────────────────────────────────────────
    const fileBytes: Buffer = Buffer.from(detail.ContentBytes, 'base64');

    // ── Step 6: Sanitize filename and resolve output path ────────────────────
    const safeName = sanitizeAttachmentName(detail.Name, `attachment-${att.Id}`);
    let targetPath: string;
    if (opts.overwrite) {
      targetPath = path.join(outputDir, safeName);
    } else {
      targetPath = deduplicateFilename(outputDir, safeName);
    }

    // ── Step 7: Atomic write ──────────────────────────────────────────────────
    await atomicWrite(targetPath, fileBytes, { mode: 0o600, overwrite: opts.overwrite });

    downloaded.push({
      id: att.Id,
      originalName: detail.Name,
      savedAs: path.basename(targetPath),
      path: targetPath,
      size: fileBytes.length,
      contentType: detail.ContentType,
      isInline: detail.IsInline,
    });
  }

  return { downloaded, skipped };
}
```

---

## 7. Error Mapping to CLI Exit Codes

The following table maps every HTTP error that can arise from the attachments endpoints to the exit code taxonomy defined in `investigation-outlook-cli.md §4.9`.

| HTTP Status | Scenario | Exit Code | Error Class | Handling |
|---|---|---|---|---|
| 401 (first attempt) | Token expired | — (retry) | — | Trigger re-auth once; on second 401, exit 4. |
| 401 (second attempt) | Re-auth failed | 4 | `AuthError` | Surface message: "authentication failed after retry". |
| 403 on list | No `Mail.Read` scope / IRM policy / mailbox access denied | 5 | `UpstreamError` | Do not retry. Include `messageId` in error. |
| 403 on detail | Item-level IRM protection (individual attachment locked) | — (skip) | — | Add to `skipped[]` with `reason: "access-denied"`. Do not abort the whole download. |
| 404 on list | `messageId` does not exist or was deleted | 5 | `UpstreamError` | Abort: the entire command target is invalid. |
| 404 on detail | Attachment was deleted between list GET and detail GET | — (skip) | — | Add to `skipped[]` with `reason: "not-found"`. Continue loop. |
| 429 | Rate limited | 5 | `UpstreamError` | Include `Retry-After` header value in error message. No auto-retry in this iteration. |
| 5xx | Server error | 5 | `UpstreamError` | No auto-retry. Surface `httpStatus` and `requestId` (from `request-id` response header). |
| Network / DNS / TLS | Connectivity failure | 5 | `UpstreamError` | Wrap the underlying Node `Error` message. Do not include the Bearer token in the error object. |
| `AbortError` | HTTP timeout | 5 | `UpstreamError` | Message: `"HTTP timeout after ${OUTLOOK_CLI_HTTP_TIMEOUT_MS}ms"`. |
| Write error (`ENOSPC`, `EACCES`, etc.) | Disk full or permissions | 6 | `IoError` | Surface `errno`, `path`. |
| Target file exists and `--overwrite` not set | Collision detected | 6 | `IoError` | Message names the offending file path. Abort before writing any attachment in the batch. |

### 7.1 Distinguishing 403 on list vs 403 on detail

A 403 on `GET /me/messages/{id}/attachments` (the list call) means the caller cannot read attachments for this message at all — likely an IRM classification on the entire message or a mailbox-level access restriction. This is fatal for the command: exit 5.

A 403 on `GET /me/messages/{id}/attachments/{attId}` (the detail call) means the individual attachment is restricted (common for `.eml` items with different IRM labels) while other attachments on the same message may be accessible. This is per-attachment: add to `skipped[]` and continue the loop.

### 7.2 Error object shape

Every `UpstreamError` and `IoError` must include:

```typescript
interface OutlookCliError {
  code: string;          // e.g. "UPSTREAM_HTTP_403", "IO_WRITE_ENOSPC"
  exitCode: number;      // 4, 5, or 6
  message: string;       // Human-readable, safe to print to stderr
  httpStatus?: number;   // Present for upstream HTTP errors
  requestId?: string;    // The "request-id" response header from Outlook, when available
  cause?: Error;         // The original Error, for debugging; Bearer/cookies MUST NOT appear here
}
```

---

## 8. MIME-Sniff and ContentType Concerns

### 8.1 Do not rely on `ContentType` for extension inference

`ContentType` on `FileAttachment` is set by the sender's mail client and is frequently wrong or generic (`application/octet-stream`). Do not rename the file or change its extension based on `ContentType`. Use `Name` as given (after sanitization).

### 8.2 Do not execute downloaded files

Write mode `0o600` prevents other users from reading or executing the file. However, on macOS, files written to `~/Downloads` or similar paths may be quarantined by Gatekeeper automatically — this is acceptable behavior and does not need to be suppressed.

### 8.3 Content-Disposition for CLI output

The CLI writes bytes to disk directly via `atomicWrite`. It does not set HTTP `Content-Disposition` headers. The sanitized `Name` is the filename on disk. No MIME interpretation occurs.

---

## 9. Assumptions and Scope

### What was assumed

| Assumption | Confidence | Impact if Wrong |
|---|---|---|
| `outlook.office.com/api/v2.0` still responds to the `GET /me/messages/{id}/attachments` endpoint despite the documented March 2024 decommission. | MEDIUM | If the endpoint is silently dead, all download commands will receive 404 or 410; the CLI must surface this clearly. Migration to Graph would be required. |
| The `@odata.type` discriminator strings (`#Microsoft.OutlookServices.*`) are exactly as documented in the v2.0 spec. | HIGH | A mismatch would cause all attachments to fall into the "unknown type" skip branch; easily debugged by logging the raw `@odata.type` value. |
| `ContentBytes` is null or absent on the list endpoint for attachments above ~3 MB and is reliable on the detail endpoint up to the same limit. | HIGH | Based on corroborating evidence from Graph API documentation and community reports. The v2.0 spec does not explicitly document this behavior. |
| `$value` is not supported on `outlook.office.com/api/v2.0` for attachments. | MEDIUM | If `$value` does work, it would enable streaming large attachments without the 4 MB base64 ceiling. This should be tested empirically on the live endpoint. |
| The CLI will not attempt to reconstruct `.eml` from `ItemAttachment`. | HIGH | In-scope decision. Would require `/$value` on Graph or an MHTML assembly step. |
| `ReferenceAttachment` never has bytes to download. | HIGH | By design in the v2.0 spec. `SourceUrl` is the only actionable field. |

### What is explicitly out of scope

- Downloading `ItemAttachment` items as `.eml` or MIME blobs.
- Following `SourceUrl` from `ReferenceAttachment` to download the referenced cloud file.
- Uploading or creating attachments.
- Handling encrypted S/MIME or IRM-protected attachment decryption.
- Streaming chunked downloads for files > 3 MB (no range-request API on v2.0).
- Pagination of the attachments collection (rare; messages typically have far fewer than the default page size of 10 attachments).

### Clarifying questions for follow-up

1. Should `$value` be empirically tested against the live `outlook.office.com/api/v2.0` endpoint using a real token? If it works, it would remove the 3 MB `ContentBytes` ceiling entirely.
2. Is there a maximum number of attachments per message that the list endpoint will page? (The v2.0 default page size for collections is typically 10; a message with 11+ attachments may require `@odata.nextLink` handling.)
3. Should the CLI surface a `--max-size-bytes` flag to skip large attachments proactively rather than discovering `ContentBytes: null` at runtime?
4. Should `ItemAttachment` ever be supported as an `.eml` export in a future iteration? If so, the Graph `/$value` path is the correct implementation route.

---

## References

| # | Source | URL | Information Gathered |
|---|---|---|---|
| 1 | Microsoft Docs — Outlook Mail REST v2.0 | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations | Endpoint paths, attachment list/detail semantics, sample requests |
| 2 | Microsoft Docs — v2.0 Complex Types Reference | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/complex-types-for-mail-contacts-calendar | Base `Attachment` resource fields: `ContentType`, `IsInline`, `LastModifiedDateTime`, `Name`, `Size` |
| 3 | Microsoft Graph — Get Attachment | https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0 | `$value` endpoint semantics, raw binary vs base64, MIME content for ItemAttachment; `@odata.type` Graph equivalents |
| 4 | Microsoft Graph — List Attachments | https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0 | Confirms `contentBytes` appears in list response; JSON sample |
| 5 | Microsoft Graph — Attach Large Files | https://learn.microsoft.com/en-us/graph/outlook-large-attachments | 4 MB REST ceiling, 3 MB binary threshold, upload-session mechanism (write-path only) |
| 6 | Microsoft Docs — ReferenceAttachment fields | https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/task-rest-operations | `SourceUrl`, `ProviderType`, `Permission`, `IsFolder`, `ThumbnailUrl`, `PreviewUrl` shape |
| 7 | PortSwigger Web Security Academy — Path Traversal | https://portswigger.net/web-security/file-path-traversal | Path traversal attack patterns and encoding bypass techniques |
| 8 | Node.js Path Traversal Security Guide | https://nodejsdesignpatterns.com/blog/nodejs-path-traversal-security/ | `path.resolve()` + boundary check pattern, `path.sep` usage |
| 9 | HackerOne — Preventing Directory Traversal | https://www.hackerone.com/blog/preventing-directory-traversal-attacks-techniques-and-tips-secure-file-access | Defense checklist: reserved names, UNC paths, drive letters |
| 10 | Xygeni — Path Traversal in File Uploads | https://xygeni.io/blog/path-traversal-in-file-uploads-how-developers-create-their-own-exploits/ | Windows reserved names list, `path.relative()` alternative check |
| 11 | CVE-2025-23084 / Node.js Windows path vuln | https://security.snyk.io/vuln/SNYK-UPSTREAM-NODE-8651420 | Windows drive-letter path traversal in Node.js; patch in v20.19.4+, v22.17.1+, v24.4.1+ |
| 12 | MS Q&A — ContentBytes missing from list | https://learn.microsoft.com/en-us/answers/questions/1080476/get-email-attachments-via-graph-api-missing-conten | Community confirmation that `contentBytes` may be absent from list endpoint |
| 13 | MS Graph Docs — fileAttachment resource | https://learn.microsoft.com/en-us/graph/api/resources/fileattachment?view=graph-rest-1.0 | `contentBytes`, `contentId`, `contentLocation`, `isInline` field definitions |

### Recommended for Deep Reading

- **Reference 3 (Graph — Get Attachment):** The `/$value` section and the per-attachment-type MIME format table are essential if `ItemAttachment` export is ever added to scope.
- **Reference 5 (Large Attachments):** If a future iteration needs to handle files > 3 MB, this documents the upload-session mechanism. The download side has no equivalent v2.0 API; Graph's `/$value` would be needed.
- **Reference 11 (CVE-2025-23084):** Ensure the Node.js runtime used in production is on a patched version before shipping the download-attachments command on Windows.
