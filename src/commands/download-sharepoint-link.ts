// src/commands/download-sharepoint-link.ts
//
// Fetch a SharePoint URL surfaced by ReferenceAttachment metadata in
// Outlook messages. Uses the secondary session captured during
// `outlook-cli login --sharepoint-host`.

import * as fs from 'node:fs';
import * as path from 'node:path';

import { ConfigurationError, OutlookCliError } from '../config/errors';
import { SharepointClient, SharepointHttpError } from '../http/sharepoint-client';
import {
  loadSharepointSession,
  defaultSharepointSessionPath,
  SharepointSession,
} from '../session/sharepoint-schema';
import { atomicWriteBuffer } from '../util/fs-atomic';
import { assertWithinDir, deduplicateFilename, sanitizeAttachmentName } from '../util/filename';

import { UsageError } from './list-mail';

/** Raised when sharepoint-session.json is missing or expired. */
export class SharepointSessionMissingError extends OutlookCliError {
  public readonly code: string = 'SHAREPOINT_SESSION_MISSING';
  public readonly exitCode: number = 4; // auth failure exit code
}

export interface DownloadSharepointLinkDeps {
  /** Override path for the SharePoint session file (test seam). */
  sharepointSessionPath?: string;
  /** HTTP timeout in milliseconds. */
  httpTimeoutMs: number;
  /** Test-seam factory; production uses default SharepointClient. */
  createSharepointClient?: (session: SharepointSession, timeoutMs: number) => SharepointClient;
}

export interface DownloadSharepointLinkOptions {
  out?: string;
  overwrite?: boolean;
}

export interface SavedRecord {
  url: string;
  name: string;
  path: string;
  size: number;
}

export type SkippedReason = 'not-found' | 'access-denied' | 'auth-required' | 'http-error';

export interface SkippedRecord {
  url: string;
  reason: SkippedReason;
  status?: number;
}

export interface DownloadSharepointLinkResult {
  outDir: string;
  saved: SavedRecord[];
  skipped: SkippedRecord[];
}

function deriveFilenameFromUrl(url: string): string {
  try {
    const u = new URL(url);
    const last = u.pathname
      .split('/')
      .filter((p) => p.length > 0)
      .pop();
    if (last && /\.[a-z0-9]+$/i.test(last)) return decodeURIComponent(last);
  } catch {
    /* ignore — fallback below */
  }
  return 'sharepoint-download';
}

function defaultClientFactory(session: SharepointSession, timeoutMs: number): SharepointClient {
  return new SharepointClient({
    bearer: session.bearer,
    cookies: session.cookies,
    timeoutMs,
  });
}

function isSessionExpired(session: SharepointSession): boolean {
  try {
    return new Date(session.tokenExpiresAt).getTime() <= Date.now();
  } catch {
    return true;
  }
}

export async function run(
  deps: DownloadSharepointLinkDeps,
  url: string,
  opts: DownloadSharepointLinkOptions,
): Promise<DownloadSharepointLinkResult> {
  if (typeof url !== 'string' || url.length === 0) {
    throw new UsageError('download-sharepoint-link: <url> is required');
  }
  if (typeof opts.out !== 'string' || opts.out.length === 0) {
    throw new ConfigurationError('download-sharepoint-link.out', ['--out flag']);
  }

  const sessionPath = deps.sharepointSessionPath ?? defaultSharepointSessionPath();
  const session = await loadSharepointSession(sessionPath);
  if (!session) {
    throw new SharepointSessionMissingError(
      `No SharePoint session found at ${sessionPath}. Run \`outlook-cli login --sharepoint-host <tenant>.sharepoint.com\` first.`,
    );
  }
  if (isSessionExpired(session)) {
    throw new SharepointSessionMissingError(
      `SharePoint session expired (tokenExpiresAt=${session.tokenExpiresAt}). Re-run \`outlook-cli login --sharepoint-host ${session.host}\`.`,
    );
  }

  const outDir = path.resolve(opts.out);
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true, mode: 0o700 });
  }

  const factory = deps.createSharepointClient ?? defaultClientFactory;
  const client = factory(session, deps.httpTimeoutMs);

  const saved: SavedRecord[] = [];
  const skipped: SkippedRecord[] = [];

  try {
    const result = await client.getBinary(url);
    const desiredName = sanitizeAttachmentName(result.filename ?? deriveFilenameFromUrl(url));
    let finalName = desiredName;
    if (!opts.overwrite) {
      const existing = new Set(fs.existsSync(outDir) ? fs.readdirSync(outDir) : []);
      finalName = deduplicateFilename(desiredName, existing);
    }
    const finalPath = path.join(outDir, finalName);
    assertWithinDir(outDir, finalPath);
    await atomicWriteBuffer(finalPath, result.bytes);
    saved.push({
      url,
      name: path.basename(finalPath),
      path: finalPath,
      size: result.size,
    });
  } catch (err) {
    if (err instanceof SharepointHttpError) {
      const reason: SkippedReason =
        err.status === 404 || err.status === 410
          ? 'not-found'
          : err.status === 403
            ? 'access-denied'
            : err.status === 401
              ? 'auth-required'
              : 'http-error';
      skipped.push({ url, reason, status: err.status });
    } else {
      throw err;
    }
  }

  return { outDir, saved, skipped };
}
