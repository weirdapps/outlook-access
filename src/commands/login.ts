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
  createClient: (s: SessionFile) => OutlookClient;
}

export interface LoginOptions {
  force?: boolean;
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
}

function buildLockPath(sessionPath: string): string {
  return path.join(path.dirname(sessionPath), '.browser.lock');
}

function toResult(session: SessionFile, sessionFile: string): LoginResult {
  return {
    status: 'ok',
    sessionFile,
    tokenExpiresAt: session.bearer.expiresAt,
    account: {
      upn: session.account.upn,
      puid: session.account.puid,
      tenantId: session.account.tenantId,
    },
  };
}

export async function run(
  deps: LoginDeps,
  opts: LoginOptions = {},
): Promise<LoginResult> {
  const force = opts.force === true;

  // Step 2: reuse cached session when allowed.
  if (!force) {
    const cached = await deps.loadSession(deps.sessionPath);
    if (cached !== null && !isExpired(cached)) {
      return toResult(cached, path.resolve(deps.sessionPath));
    }
  }

  // Step 1 (reordered): acquire the single-writer lock before opening Chrome.
  const lockPath = buildLockPath(deps.sessionPath);
  const release = await acquireLock(lockPath);

  try {
    // Step 3 + 4: run the Playwright capture and build the SessionFile.
    // doAuthCapture is injected by cli.ts; it wraps captureOutlookSession and
    // attaches `version` + `capturedAt` to produce a fully-formed SessionFile.
    const session = await deps.doAuthCapture();

    // Step 5: persist atomically (mode 0600).
    await deps.saveSession(deps.sessionPath, session);

    // Step 7: return the response shape. (Step 6 — lock release — happens in
    // the `finally` block below.)
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
