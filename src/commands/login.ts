// src/commands/login.ts
//
// Open Chrome via Playwright, capture a fresh Outlook session, persist it.
// See project-design.md §2.13.1.

import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import type { OutlookClient } from '../http/outlook-client';
import type { SessionFile } from '../session/schema';
import { isExpired } from '../session/store';
import { acquireLock } from '../auth/lock';

export interface LoginDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  /** Optional — required only when LoginOptions.sharepointHost is set. */
  doAuthCaptureWithSharepoint?: (host: string) => Promise<{
    session: SessionFile;
    sharepointPath: string;
  }>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface LoginOptions {
  force?: boolean;
  /** When set, also capture a SharePoint session for this host
   *  (e.g. "nbg.sharepoint.com") into ~/.outlook-cli/sharepoint-session.json. */
  sharepointHost?: string;
}

export interface LoginResult {
  status: 'ok';
  sessionFile: string;
  tokenExpiresAt: string;
  account: {
    upn: string;
    puid: string;
    tenantId: string;
  };
  /** Path to the persisted SharePoint session file when --sharepoint-host was set. */
  sharepointSessionFile?: string;
}

function buildLockPath(sessionPath: string): string {
  return path.join(path.dirname(sessionPath), '.browser.lock');
}

function toResult(
  session: SessionFile,
  sessionFile: string,
  sharepointSessionFile?: string,
): LoginResult {
  return {
    status: 'ok',
    sessionFile,
    tokenExpiresAt: session.bearer.expiresAt,
    account: {
      upn: session.account.upn,
      puid: session.account.puid,
      tenantId: session.account.tenantId,
    },
    ...(sharepointSessionFile ? { sharepointSessionFile } : {}),
  };
}

export async function run(deps: LoginDeps, opts: LoginOptions = {}): Promise<LoginResult> {
  const force = opts.force === true;

  const sharepointHost = opts.sharepointHost;

  // Step 2: reuse cached session when allowed (skip when --sharepoint-host —
  // we always need to open the browser for the SharePoint capture leg).
  if (!force && !sharepointHost) {
    const cached = await deps.loadSession(deps.sessionPath);
    if (cached !== null && !isExpired(cached)) {
      return toResult(cached, path.resolve(deps.sessionPath));
    }
  }

  // Step 1 (reordered): acquire the single-writer lock before opening Chrome.
  const lockPath = buildLockPath(deps.sessionPath);
  const release = await acquireLock(lockPath);

  try {
    if (sharepointHost) {
      if (!deps.doAuthCaptureWithSharepoint) {
        throw new Error(
          'login: --sharepoint-host was set but doAuthCaptureWithSharepoint is not wired',
        );
      }
      const { session, sharepointPath } = await deps.doAuthCaptureWithSharepoint(sharepointHost);
      // saveSession is also called inside doAuthCaptureWithSharepoint, but we
      // re-call it here to keep the contract symmetric with the non-SharePoint
      // path. Atomic write is idempotent.
      await deps.saveSession(deps.sessionPath, session);
      return toResult(session, path.resolve(deps.sessionPath), sharepointPath);
    }

    // Standard Outlook-only path
    const session = await deps.doAuthCapture();
    await deps.saveSession(deps.sessionPath, session);
    return toResult(session, path.resolve(deps.sessionPath));
  } finally {
    try {
      await release();
    } catch {
      // Lock release failure is non-fatal: the lock file uses a PID sentinel
      // and a stale entry will be reclaimed on the next invocation.
    }
  }
}
