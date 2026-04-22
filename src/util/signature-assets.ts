// src/util/signature-assets.ts
//
// Inline-image support for the captured signature.
//
// Signatures are HTML — but Outlook signatures often reference embedded
// images via `<img src="cid:CONTENT_ID">`. For those images to render in
// sent mail, the message must include a matching FileAttachment with
// `IsInline: true` and `ContentId: CONTENT_ID`.
//
// This module:
//   1. Saves attachments to `<assetsDir>/<sanitized-cid>.bin` + manifest.json
//      (used by `capture-signature` after fetching attachments).
//   2. Loads attachments back as `SendFileAttachment[]` filtered by the
//      `cid:` references actually present in the signature HTML
//      (used by `send-mail` and reply/forward).

import { promises as fs } from 'node:fs';
import * as path from 'node:path';

import type { SendFileAttachment } from '../http/outlook-client';

/** Filename that holds the contentId → asset metadata mapping. */
const MANIFEST_NAME = 'manifest.json';

/** Manifest entry for a single inline asset. */
export interface SignatureAssetEntry {
  /** Original Outlook ContentId (unsanitized). */
  contentId: string;
  /** Sanitized filename used on disk (no path separators). */
  fileName: string;
  /** Mime type (e.g., image/png). */
  contentType: string;
  /** Original display name from the source message attachment. */
  originalName: string;
}

/** Persisted manifest shape. */
export interface SignatureAssetsManifest {
  version: 1;
  capturedAt: string;
  /** Source message id from which the signature + assets were extracted. */
  sourceMessageId: string;
  assets: SignatureAssetEntry[];
}

/**
 * Save signature attachments to disk. Writes one binary file per asset and
 * a single manifest.json. Returns the manifest written.
 *
 * Caller-provided sanitization of contentId for the filename: replace
 * anything that isn't [A-Za-z0-9._@-] with `_`.
 */
export async function saveSignatureAssets(opts: {
  assetsDir: string;
  sourceMessageId: string;
  attachments: Array<{
    contentId: string;
    contentType: string;
    contentBytesBase64: string;
    name: string;
  }>;
  /** Override for tests. */
  writeFile?: (p: string, data: Buffer | string) => Promise<void>;
  mkdir?: (p: string, opts: { recursive: boolean; mode: number }) => Promise<unknown>;
}): Promise<SignatureAssetsManifest> {
  const writer = opts.writeFile ?? ((p, d) => fs.writeFile(p, d, { mode: 0o600 }));
  const mkdir = opts.mkdir ?? ((p, o) => fs.mkdir(p, o));
  await mkdir(opts.assetsDir, { recursive: true, mode: 0o700 });

  const entries: SignatureAssetEntry[] = [];
  for (const att of opts.attachments) {
    const fileName = sanitizeContentIdForFile(att.contentId);
    const filePath = path.join(opts.assetsDir, fileName);
    await writer(filePath, Buffer.from(att.contentBytesBase64, 'base64'));
    entries.push({
      contentId: att.contentId,
      fileName,
      contentType: att.contentType,
      originalName: att.name,
    });
  }
  const manifest: SignatureAssetsManifest = {
    version: 1,
    capturedAt: new Date().toISOString(),
    sourceMessageId: opts.sourceMessageId,
    assets: entries,
  };
  await writer(path.join(opts.assetsDir, MANIFEST_NAME), JSON.stringify(manifest, null, 2));
  return manifest;
}

/**
 * Load the manifest from disk. Returns `null` if the manifest is missing
 * or malformed (signature has no inline images, or never captured).
 */
export async function loadManifest(
  assetsDir: string,
  reader?: (p: string) => Promise<Buffer>,
): Promise<SignatureAssetsManifest | null> {
  const r = reader ?? ((p: string) => fs.readFile(p));
  try {
    const buf = await r(path.join(assetsDir, MANIFEST_NAME));
    const parsed = JSON.parse(buf.toString('utf-8'));
    if (parsed && parsed.version === 1 && Array.isArray(parsed.assets)) {
      return parsed as SignatureAssetsManifest;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Scan the signature HTML for `<img src="cid:XXX">` references, look up
 * matching assets in the manifest, load their bytes, and return a list of
 * `SendFileAttachment` ready to splice into the SendMailPayload.
 *
 * Returns `[]` if there are no cid refs OR the manifest is missing.
 *
 * Refs found in HTML but with no matching manifest entry are silently
 * skipped — the email will still send, but the broken image will show in
 * Outlook. The `unmatchedRefs` field reports them to the caller so it can
 * surface a warning if desired.
 */
export async function loadSignatureAttachments(opts: {
  signatureHtml: string;
  assetsDir: string;
  reader?: (p: string) => Promise<Buffer>;
}): Promise<{
  attachments: SendFileAttachment[];
  unmatchedRefs: string[];
}> {
  const refs = extractCidReferences(opts.signatureHtml);
  if (refs.length === 0) {
    return { attachments: [], unmatchedRefs: [] };
  }
  const manifest = await loadManifest(opts.assetsDir, opts.reader);
  if (!manifest) {
    // No manifest = all refs are unmatched.
    return { attachments: [], unmatchedRefs: refs };
  }
  const r = opts.reader ?? ((p: string) => fs.readFile(p));
  const byContentId = new Map(manifest.assets.map((a) => [a.contentId, a]));

  const attachments: SendFileAttachment[] = [];
  const unmatchedRefs: string[] = [];
  for (const cid of refs) {
    const entry = byContentId.get(cid);
    if (!entry) {
      unmatchedRefs.push(cid);
      continue;
    }
    let buf: Buffer;
    try {
      buf = await r(path.join(opts.assetsDir, entry.fileName));
    } catch {
      unmatchedRefs.push(cid);
      continue;
    }
    attachments.push({
      '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
      Name: entry.originalName,
      ContentType: entry.contentType,
      ContentBytes: buf.toString('base64'),
      IsInline: true,
      ContentId: entry.contentId,
      Size: buf.length,
    });
  }
  return { attachments, unmatchedRefs };
}

/**
 * Find every `src="cid:XXX"` in the HTML and return the unique XXX values
 * in document order. Quote style is normalized — accepts single or double.
 */
export function extractCidReferences(html: string): string[] {
  const re = /src\s*=\s*["']cid:([^"']+)["']/gi;
  const seen = new Set<string>();
  const out: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) {
    const cid = m[1]!;
    if (!seen.has(cid)) {
      seen.add(cid);
      out.push(cid);
    }
  }
  return out;
}

/**
 * Sanitize a contentId for use as a filename. Outlook contentIds typically
 * look like `image001.png@01DCD27B.DECD9E60` which has `.` and `@` — both
 * are filesystem-safe. Anything outside [A-Za-z0-9._@-] becomes `_`.
 */
export function sanitizeContentIdForFile(contentId: string): string {
  return contentId.replace(/[^A-Za-z0-9._@-]/g, '_');
}
