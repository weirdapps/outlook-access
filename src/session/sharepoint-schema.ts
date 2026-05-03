// src/session/sharepoint-schema.ts
//
// Schema + IO helpers for ~/.outlook-cli/sharepoint-session.json — the
// secondary session file captured during `outlook-cli login --sharepoint-host`.
// Used by download-sharepoint-link to fetch ReferenceAttachment URLs that
// resolve to SharePoint or OneDrive-for-Business hosts.

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

export interface SharepointSession {
  version: 1;
  /** SharePoint host, e.g. "nbg.sharepoint.com". */
  host: string;
  /** Bearer token (no "Bearer " prefix). */
  bearer: string;
  /** Serialized cookie header value, e.g. "rtFa=...; FedAuth=...". */
  cookies: string;
  /** ISO-8601 UTC timestamp of capture. */
  capturedAt: string;
  /** ISO-8601 UTC, derived from JWT exp. */
  tokenExpiresAt: string;
}

export class SharepointSessionParseError extends Error {
  constructor(msg: string) {
    super(msg);
    this.name = 'SharepointSessionParseError';
  }
}

export function defaultSharepointSessionPath(): string {
  return path.join(os.homedir(), '.outlook-cli', 'sharepoint-session.json');
}

export function parseSharepointSession(json: string): SharepointSession {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch (e) {
    throw new SharepointSessionParseError(`Invalid JSON: ${(e as Error).message}`);
  }
  if (typeof raw !== 'object' || raw === null) {
    throw new SharepointSessionParseError('Expected JSON object');
  }
  const obj = raw as Record<string, unknown>;
  if (obj.version !== 1) {
    throw new SharepointSessionParseError(`Unsupported version: ${String(obj.version)}`);
  }
  for (const key of ['host', 'bearer', 'cookies', 'capturedAt', 'tokenExpiresAt']) {
    if (typeof obj[key] !== 'string' || (obj[key] as string).length === 0) {
      throw new SharepointSessionParseError(`Missing or invalid "${key}"`);
    }
  }
  return obj as unknown as SharepointSession;
}

export function serializeSharepointSession(s: SharepointSession): string {
  return JSON.stringify(s, null, 2);
}

export async function loadSharepointSession(filePath: string): Promise<SharepointSession | null> {
  try {
    const data = await fs.promises.readFile(filePath, 'utf8');
    return parseSharepointSession(data);
  } catch (err: unknown) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return null;
    throw err;
  }
}

export async function saveSharepointSession(
  filePath: string,
  session: SharepointSession,
): Promise<void> {
  const dir = path.dirname(filePath);
  await fs.promises.mkdir(dir, { recursive: true, mode: 0o700 });
  try {
    await fs.promises.chmod(dir, 0o700);
  } catch {
    /* tolerate on existing dirs */
  }
  const tmp = filePath + '.tmp';
  await fs.promises.writeFile(tmp, serializeSharepointSession(session), {
    mode: 0o600,
  });
  await fs.promises.rename(tmp, filePath);
}
