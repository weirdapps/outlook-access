// src/commands/auth-renew.ts
//
// Silent (headless) bearer renewal. Uses the persisted Playwright profile
// to re-issue an OWA bearer without opening a visible browser window.
//
// Works while the device-trust cookie (ESTSAUTHPERSISTENT, ~90 days) is
// alive. When that cookie expires or NBG forces re-MFA, this command fails
// with AuthError(AUTH_LOGIN_TIMEOUT) and the caller must run `login`.

import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import type { SessionFile } from '../session/schema';
import { AuthError } from '../config/errors';
import { acquireLock } from '../auth/lock';
import { captureOutlookSession, AuthCaptureError } from '../auth/browser-capture';

/** Default headless renewal timeout — much shorter than interactive login. */
const DEFAULT_RENEW_TIMEOUT_MS = 30_000;

export interface AuthRenewDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (p: string) => Promise<SessionFile | null>;
  saveSession: (p: string, s: SessionFile) => Promise<void>;
}

export interface AuthRenewOptions {
  /** Override the renew-specific timeout (default 30000ms). */
  timeoutMs?: number;
}

export interface AuthRenewResult {
  status: 'ok';
  sessionFile: string;
  tokenExpiresAt: string;
  account: { upn: string; puid: string; tenantId: string };
  /** Wall-clock duration of the renewal in milliseconds. */
  durationMs: number;
}

function buildLockPath(sessionPath: string): string {
  return path.join(path.dirname(sessionPath), '.browser.lock');
}

export async function run(
  deps: AuthRenewDeps,
  opts: AuthRenewOptions = {},
): Promise<AuthRenewResult> {
  const timeoutMs = opts.timeoutMs ?? DEFAULT_RENEW_TIMEOUT_MS;

  // A renewal only makes sense if a prior interactive login left a profile
  // behind. Fail fast otherwise — the caller should run `login`.
  const existing = await deps.loadSession(deps.sessionPath);
  if (existing === null) {
    throw new AuthError(
      'AUTH_NO_REAUTH',
      'No cached session to renew. Run `outlook-cli login` first.',
    );
  }

  const release = await acquireLock(buildLockPath(deps.sessionPath));
  const t0 = Date.now();

  try {
    const captured = await captureOutlookSession({
      profileDir: deps.config.profileDir,
      chromeChannel: deps.config.chromeChannel,
      loginTimeoutMs: timeoutMs,
      headless: true,
    });

    const session: SessionFile = {
      version: 1,
      capturedAt: new Date().toISOString(),
      account: captured.account,
      bearer: captured.bearer,
      cookies: captured.cookies,
      anchorMailbox: captured.anchorMailbox,
    };
    await deps.saveSession(deps.sessionPath, session);

    return {
      status: 'ok',
      sessionFile: path.resolve(deps.sessionPath),
      tokenExpiresAt: session.bearer.expiresAt,
      account: session.account,
      durationMs: Date.now() - t0,
    };
  } catch (err) {
    if (err instanceof AuthCaptureError) {
      // Either the device-trust cookie is gone, NBG forced re-MFA, or some
      // navigation glitch. In all cases, the path forward is interactive.
      throw new AuthError(
        'AUTH_LOGIN_TIMEOUT',
        `Headless renewal failed (${err.code}): ${err.message}. Run \`outlook-cli login\`.`,
        err,
      );
    }
    throw err;
  } finally {
    try {
      await release();
    } catch {
      // Stale lock entries are reclaimed on next invocation.
    }
  }
}
