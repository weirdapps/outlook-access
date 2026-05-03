// src/commands/download-attachments.ts
//
// Save all (non-inline by default) FileAttachment content bytes from a
// message into a user-chosen output directory.
// See project-design.md §2.13.5 and research doc §6.

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import { ConfigurationError, IoError } from '../config/errors';
import type { OutlookClient } from '../http/outlook-client';
import { ApiError } from '../http/errors';
import type {
  Attachment,
  AttachmentEnvelope,
  FileAttachment,
  ODataListResponse,
  ReferenceAttachment,
} from '../http/types';
import type { SessionFile } from '../session/schema';
import { atomicWriteBuffer } from '../util/fs-atomic';
import { assertWithinDir, deduplicateFilename, sanitizeAttachmentName } from '../util/filename';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface DownloadAttachmentsDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface DownloadAttachmentsOptions {
  out?: string;
  overwrite?: boolean;
  includeInline?: boolean;
}

export interface SavedRecord {
  id: string;
  name: string;
  path: string;
  size: number;
}

export type SkippedReason =
  | 'inline'
  | 'reference-attachment'
  | 'item-attachment'
  | 'unknown-attachment-type'
  | 'content-bytes-null'
  | 'not-found'
  | 'access-denied';

export interface SkippedRecord {
  id: string;
  name: string;
  reason: SkippedReason;
  sourceUrl?: string;
  odataType?: string;
}

export interface DownloadAttachmentsResult {
  messageId: string;
  outDir: string;
  saved: SavedRecord[];
  skipped: SkippedRecord[];
}

function isFileAttachment(a: Attachment): a is FileAttachment {
  return a['@odata.type'] === '#Microsoft.OutlookServices.FileAttachment';
}

function isReferenceAttachment(a: Attachment): a is ReferenceAttachment {
  return a['@odata.type'] === '#Microsoft.OutlookServices.ReferenceAttachment';
}

function isItemAttachment(a: Attachment): boolean {
  return a['@odata.type'] === '#Microsoft.OutlookServices.ItemAttachment';
}

export async function run(
  deps: DownloadAttachmentsDeps,
  id: string,
  opts: DownloadAttachmentsOptions = {},
): Promise<DownloadAttachmentsResult> {
  if (typeof id !== 'string' || id.length === 0) {
    throw new UsageError('download-attachments: <id> is required');
  }
  if (typeof opts.out !== 'string' || opts.out.length === 0) {
    // Mandatory option not provided. Per design, this is a ConfigurationError.
    throw new ConfigurationError('download-attachments.out', ['--out flag']);
  }

  const overwrite = opts.overwrite === true;
  const includeInline = opts.includeInline === true;

  const outDir = path.resolve(opts.out);
  await ensureOutDir(outDir);

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const encodedId = encodeURIComponent(id);

  let listing: ODataListResponse<AttachmentEnvelope>;
  try {
    listing = await client.get<ODataListResponse<AttachmentEnvelope>>(
      `/api/v2.0/me/messages/${encodedId}/attachments`,
    );
  } catch (err) {
    throw mapHttpError(err);
  }

  const saved: SavedRecord[] = [];
  const skipped: SkippedRecord[] = [];
  const existingNames = new Set<string>();

  for (const att of listing.value ?? []) {
    if (isReferenceAttachment(att)) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'reference-attachment',
        sourceUrl: att.SourceUrl,
      });
      continue;
    }
    if (isItemAttachment(att)) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'item-attachment',
      });
      continue;
    }
    if (!isFileAttachment(att)) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'unknown-attachment-type',
        odataType: att['@odata.type'],
      });
      continue;
    }

    // FileAttachment
    if (att.IsInline && !includeInline) {
      skipped.push({ id: att.Id, name: att.Name, reason: 'inline' });
      continue;
    }

    // Per-item detail fetch to get ContentBytes reliably.
    let detail: FileAttachment;
    try {
      detail = await client.get<FileAttachment>(
        `/api/v2.0/me/messages/${encodedId}/attachments/${encodeURIComponent(att.Id)}`,
      );
    } catch (err) {
      if (err instanceof ApiError && err.httpStatus === 404) {
        skipped.push({ id: att.Id, name: att.Name, reason: 'not-found' });
        continue;
      }
      if (err instanceof ApiError && err.httpStatus === 403) {
        skipped.push({ id: att.Id, name: att.Name, reason: 'access-denied' });
        continue;
      }
      throw mapHttpError(err);
    }

    if (detail.ContentBytes === null || detail.ContentBytes === undefined) {
      skipped.push({
        id: att.Id,
        name: att.Name,
        reason: 'content-bytes-null',
      });
      continue;
    }

    const sanitized = sanitizeAttachmentName(detail.Name ?? att.Name ?? '');
    const deduped = deduplicateFilename(sanitized, existingNames);
    const targetPath = assertWithinDir(outDir, deduped);

    let buf: Buffer;
    try {
      buf = Buffer.from(detail.ContentBytes, 'base64');
    } catch (err) {
      throw new IoError({
        code: 'IO_ATTACHMENT_DECODE',
        message: `Failed to decode ContentBytes for attachment ${att.Id}.`,
        cause: err,
      });
    }

    // Atomic write with overwrite guard per design §2.13.5 step 8.
    // `atomicWriteBuffer` uses O_CREAT|O_EXCL on the temp file and, when
    // `overwrite` is false, throws IoError('IO_WRITE_EEXIST') if the final
    // path already exists. The explicit 0o644 mode matches typical
    // user-visible output files (not the session file's 0o600).
    await atomicWriteBuffer(targetPath, buf, {
      mode: 0o644,
      overwrite,
    });

    existingNames.add(deduped);
    saved.push({
      id: att.Id,
      name: deduped,
      path: targetPath,
      size: buf.byteLength,
    });
  }

  return { messageId: id, outDir, saved, skipped };
}

async function ensureOutDir(outDir: string): Promise<void> {
  // User-chosen output dir: do NOT force 0o700. Design §2.13.5 step 1 says
  // "mkdir recursive (mode 0700 is NOT set here — user-chosen output dir
  // uses default umask)."
  try {
    await fs.promises.mkdir(outDir, { recursive: true });
  } catch (err) {
    throw new IoError({
      code: 'IO_MKDIR_EACCES',
      message: `Cannot create output directory: ${outDir}`,
      path: outDir,
      cause: err,
    });
  }
}
