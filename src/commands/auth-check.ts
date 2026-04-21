// src/commands/auth-check.ts
//
// Verify that the cached session is present, non-expired, and currently
// accepted by Outlook. Does NOT auto-reauth (design §2.13.2).

import type { CliConfig } from '../config/config';
import type { SessionFile } from '../session/schema';
import { isExpired } from '../session/store';
import { createOutlookClient, type OutlookClient } from '../http/outlook-client';
import { AuthError, ApiError, NetworkError } from '../http/errors';
import { UpstreamError } from '../config/errors';

export interface AuthCheckDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface AuthCheckOptions {
  // No per-command options.
}

export type AuthCheckStatus = 'ok' | 'expired' | 'missing' | 'rejected';

export interface AuthCheckResult {
  status: AuthCheckStatus;
  tokenExpiresAt: string | null;
  account: { upn: string } | null;
}

// Minimal shape for the /me response we consume. Outlook v2 returns PascalCase.
interface MeResponse {
  EmailAddress?: string;
  Id?: string;
}

export async function run(
  deps: AuthCheckDeps,
  _opts: AuthCheckOptions = {},
): Promise<AuthCheckResult> {
  const session = await deps.loadSession(deps.sessionPath);

  if (session === null) {
    return { status: 'missing', tokenExpiresAt: null, account: null };
  }

  if (isExpired(session)) {
    return {
      status: 'expired',
      tokenExpiresAt: session.bearer.expiresAt,
      account: { upn: session.account.upn },
    };
  }

  // Build client with noAutoReauth:true so the 401 path does NOT call the
  // browser. The onReauthNeeded callback is required but must never run here.
  const client = createOutlookClient({
    session,
    httpTimeoutMs: deps.config.httpTimeoutMs,
    noAutoReauth: true,
    onReauthNeeded: async () => {
      // Should be unreachable because noAutoReauth is true. Surface a hard
      // error if ever invoked so the contract violation is obvious.
      throw new Error('auth-check: onReauthNeeded must not be invoked');
    },
  });

  try {
    await client.get<MeResponse>('/api/v2.0/me');
    return {
      status: 'ok',
      tokenExpiresAt: session.bearer.expiresAt,
      account: { upn: session.account.upn },
    };
  } catch (err) {
    if (err instanceof AuthError) {
      // 401: rejected credentials. Per design, auth-check exits 0 with
      // status "rejected" rather than propagating.
      return {
        status: 'rejected',
        tokenExpiresAt: session.bearer.expiresAt,
        account: { upn: session.account.upn },
      };
    }
    if (err instanceof ApiError) {
      // Non-401 upstream HTTP error → propagate as UpstreamError (exit 5).
      throw new UpstreamError({
        code: `UPSTREAM_HTTP_${err.httpStatus}`,
        message: err.message,
        httpStatus: err.httpStatus,
        requestId: err.requestId,
        url: err.url,
        cause: err,
      });
    }
    if (err instanceof NetworkError) {
      throw new UpstreamError({
        code: err.timedOut ? 'UPSTREAM_TIMEOUT' : 'UPSTREAM_NETWORK',
        message: err.message,
        url: err.url,
        cause: err,
      });
    }
    throw err;
  }
}
