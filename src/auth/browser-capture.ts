// src/auth/browser-capture.ts
//
// Unit 3 — Playwright-driven capture of the first `Authorization: Bearer <jwt>`
// header that Outlook Web sends to outlook.office.com, plus the session
// cookies needed to replay requests from Node.
//
// Design ref: docs/design/project-design.md §2.7
// Research:   docs/research/playwright-token-capture.md §3, §9 (INIT_SCRIPT_TEXT)
//             docs/design/refined-request-outlook-cli.md §6.3 (re-auth flow)

import * as fs from 'node:fs';
import { chromium, BrowserContext, Page } from 'playwright';

import { Cookie } from '../session/schema';
import type { SharepointSession } from '../session/sharepoint-schema';
import { decodeJwt, JwtClaims } from './jwt';
import { captureSharepointFromContext } from './sharepoint-capture';

// ─────────────────────────────────────────────────────────────────────────────
// Public types
// ─────────────────────────────────────────────────────────────────────────────

export interface CaptureResult {
  bearer: {
    /** Raw JWT (no "Bearer " prefix). */
    token: string;
    /** ISO8601 UTC, derived from JWT exp. */
    expiresAt: string;
    /** From JWT aud. */
    audience: string;
    /** From JWT scp split on whitespace. May be []. */
    scopes: string[];
  };
  cookies: Cookie[];
  account: {
    upn: string;
    puid: string;
    tenantId: string;
  };
  /** Pre-computed "PUID:<puid>@<tenantId>". */
  anchorMailbox: string;
  /** When CaptureOptions.sharepointHost is set, the SharePoint session
   *  captured from the same persistent context after Outlook auth. */
  sharepoint?: SharepointSession;
}

export interface CaptureOptions {
  /** Persistent Chrome profile dir. Created with mode 0700 if missing. */
  profileDir: string;
  /** Playwright channel — "chrome", "msedge", etc. */
  chromeChannel: string;
  /** Max wall-clock time waiting for first Bearer capture. */
  loginTimeoutMs: number;
  /** When true, the browser opens even if a cached profile could do silent SSO. */
  force?: boolean;
  /** When set, capture a SharePoint session from the same persistent context
   *  after Outlook auth succeeds. Host must end in ".sharepoint.com" — see
   *  captureSharepointFromContext for validation. */
  sharepointHost?: string;
  /** When true, launch Chromium headless. Requires a persistent profile with
   *  valid ESTSAUTHPERSISTENT cookie so Entra silently re-issues a bearer
   *  without user interaction. Used by `outlook-cli auth-renew`. */
  headless?: boolean;
}

export type AuthCaptureErrorCode =
  | 'AUTH_CANCELLED'
  | 'LOGIN_TIMEOUT'
  | 'NO_TOKEN';

export class AuthCaptureError extends Error {
  public readonly code: AuthCaptureErrorCode;

  constructor(code: AuthCaptureErrorCode, message: string) {
    super(message);
    this.name = 'AuthCaptureError';
    this.code = code;
    // Make the prototype chain work when targeting ES5 outputs.
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Init script (installed as a string; runs in the browser).
// Reproduced verbatim from docs/research/playwright-token-capture.md §9.
// ─────────────────────────────────────────────────────────────────────────────

export const INIT_SCRIPT_TEXT = `
(function () {
  if (window.__outlookCliHooked) return;
  window.__outlookCliHooked = true;

  let reported = false;

  const TARGET_PREFIXES = [
    'https://outlook.office.com/api/v2.0/',
    'https://outlook.office.com/ows/',
    'https://outlook.office365.com/api/v2.0/',
    'https://outlook.office365.com/ows/',
  ];

  function isTargetUrl(url) {
    const s = typeof url === 'string' ? url : (url && url.href ? url.href : String(url));
    return TARGET_PREFIXES.some(prefix => s.startsWith(prefix));
  }

  function extractBearer(headers) {
    if (!headers) return null;
    if (typeof headers.get === 'function') {
      return headers.get('authorization') || headers.get('Authorization') || null;
    }
    if (Array.isArray(headers)) {
      const pair = headers.find(([k]) => k.toLowerCase() === 'authorization');
      return pair ? pair[1] : null;
    }
    const key = Object.keys(headers).find(k => k.toLowerCase() === 'authorization');
    return key ? headers[key] : null;
  }

  function tryReport(url, authHeader) {
    if (reported) return;
    if (!authHeader || !authHeader.startsWith('Bearer ')) return;
    if (!isTargetUrl(url)) return;
    reported = true;
    window.__outlookCliReportBearer({ url: String(url), token: authHeader });
  }

  // Patch fetch
  const originalFetch = window.fetch;
  window.fetch = function fetch(input, init) {
    try {
      let url = input;
      let authHeader = null;
      if (input instanceof Request) {
        url = input.url;
        authHeader = extractBearer(input.headers);
        if (init && init.headers) {
          const override = extractBearer(init.headers);
          if (override) authHeader = override;
        }
      } else {
        url = input;
        authHeader = init && init.headers ? extractBearer(init.headers) : null;
      }
      tryReport(url, authHeader);
    } catch (_) {}
    return originalFetch.apply(this, arguments);
  };

  // Patch XMLHttpRequest
  const OriginalXHR = window.XMLHttpRequest;
  function PatchedXHR() {
    const xhr = new OriginalXHR();
    let _url = '';
    let _authHeader = null;
    const originalOpen = xhr.open.bind(xhr);
    xhr.open = function open(method, url) {
      _url = url;
      _authHeader = null;
      return originalOpen.apply(xhr, arguments);
    };
    const originalSetRequestHeader = xhr.setRequestHeader.bind(xhr);
    xhr.setRequestHeader = function setRequestHeader(name, value) {
      if (name.toLowerCase() === 'authorization') _authHeader = value;
      return originalSetRequestHeader.apply(xhr, arguments);
    };
    const originalSend = xhr.send.bind(xhr);
    xhr.send = function send() {
      try { tryReport(_url, _authHeader); } catch (_) {}
      return originalSend.apply(xhr, arguments);
    };
    return xhr;
  }
  Object.setPrototypeOf(PatchedXHR, OriginalXHR);
  Object.setPrototypeOf(PatchedXHR.prototype, OriginalXHR.prototype);
  Object.defineProperty(PatchedXHR, 'name', { value: 'XMLHttpRequest' });
  window.XMLHttpRequest = PatchedXHR;
})();
`;

// ─────────────────────────────────────────────────────────────────────────────
// Implementation
// ─────────────────────────────────────────────────────────────────────────────

/** Payload shape relayed from the in-page hook via exposeBinding. */
interface BearerPayload {
  url: string;
  token: string; // Full "Bearer eyJ..." string
}

/** Cookie domain whitelist — matches the set Outlook + Entra actually needs. */
const COOKIE_DOMAIN_SUFFIXES: readonly string[] = [
  '.office.com',
  '.outlook.office.com',
  '.outlook.office365.com',
  '.login.microsoftonline.com',
  '.microsoftonline.com',
  'office.com',
  'outlook.office.com',
  'outlook.office365.com',
  'login.microsoftonline.com',
];

/**
 * Produce a redacted preview of a bearer token for any future diagnostic
 * logging. NEVER return or log the full token value.
 * Kept exported (underscore prefix) so downstream units can reuse the exact
 * redaction contract if they wire up debug logging.
 */
export function _redactToken(token: string): string {
  if (!token) return '';
  if (token.length <= 10) return '…';
  return token.slice(0, 10) + '…';
}

export async function captureOutlookSession(
  opts: CaptureOptions,
): Promise<CaptureResult> {
  // 1. Ensure profile dir exists with mode 0o700.
  fs.mkdirSync(opts.profileDir, { recursive: true, mode: 0o700 });
  try {
    fs.chmodSync(opts.profileDir, 0o700);
  } catch {
    // Non-fatal — mkdirSync already set the mode on creation. On existing dirs
    // where we don't own the inode this may fail; we tolerate that.
  }

  // 2. Launch persistent Chrome.
  const context: BrowserContext = await chromium.launchPersistentContext(
    opts.profileDir,
    {
      channel: opts.chromeChannel,
      headless: opts.headless === true,
      viewport: { width: 1280, height: 900 },
      args: ['--no-first-run', '--no-default-browser-check'],
    },
  );

  try {
    // 3. Set up the capture promise plumbing BEFORE creating/using any pages.
    let resolveCapture!: (payload: BearerPayload) => void;
    let rejectCapture!: (err: Error) => void;
    let settled = false;

    const capturePromise = new Promise<BearerPayload>((resolve, reject) => {
      resolveCapture = (v) => {
        if (settled) return;
        settled = true;
        resolve(v);
      };
      rejectCapture = (e) => {
        if (settled) return;
        settled = true;
        reject(e);
      };
    });

    // 4. Register the exposeBinding FIRST, then the init script.
    await context.exposeBinding(
      '__outlookCliReportBearer',
      (_source, payload: unknown) => {
        // Validate the shape before acting on it.
        if (
          !payload ||
          typeof payload !== 'object' ||
          typeof (payload as BearerPayload).token !== 'string' ||
          typeof (payload as BearerPayload).url !== 'string'
        ) {
          return;
        }
        resolveCapture(payload as BearerPayload);
      },
    );

    await context.addInitScript(INIT_SCRIPT_TEXT);

    // 5. Open a fresh page. We do NOT rely on context.pages()[0] because the
    //    research doc flags a Playwright bug (#28692) where init scripts do
    //    not fire on pre-restored pages. A new page + explicit goto is the
    //    reliable path.
    const page: Page = await context.newPage();

    // 6. Close-guards — reject if the user closes the browser early.
    const onClose = (): void => {
      rejectCapture(
        new AuthCaptureError(
          'AUTH_CANCELLED',
          'user cancelled: browser closed before Bearer token was captured',
        ),
      );
    };
    page.once('close', onClose);
    context.once('close', onClose);

    // 7. Timeout guard.
    const timeoutHandle = setTimeout(() => {
      rejectCapture(
        new AuthCaptureError(
          'LOGIN_TIMEOUT',
          `login timeout: no Bearer token captured within ${opts.loginTimeoutMs}ms`,
        ),
      );
    }, opts.loginTimeoutMs);
    // Don't keep the event loop alive if nothing else is pending.
    if (typeof timeoutHandle.unref === 'function') timeoutHandle.unref();

    // 8. Navigate. CRITICAL: even if the profile has a cached tab, we MUST
    //    issue an explicit goto so the init script fires (Playwright #28692).
    //    Navigation errors are non-fatal — the SPA may still complete the auth
    //    dance, and the timeout guard covers the pathological case.
    try {
      await page.goto('https://outlook.office.com/mail/', {
        waitUntil: 'domcontentloaded',
        timeout: opts.loginTimeoutMs,
      });
    } catch {
      // Swallow: we trust the capture / timeout promises to drive the outcome.
    }

    // 9. Wait for the token.
    let payload: BearerPayload;
    try {
      payload = await capturePromise;
    } finally {
      clearTimeout(timeoutHandle);
      try { page.off('close', onClose); } catch { /* ignore */ }
      try { context.off('close', onClose); } catch { /* ignore */ }
    }

    // 10. Normalise the token: store the JWT only. Outlook occasionally emits
    //     "Bearer  <jwt>" with two (or more) whitespace chars after the
    //     scheme, so strip any run of whitespace — not a literal single space.
    const rawHeader = payload.token;
    const jwt = rawHeader.replace(/^Bearer\s+/i, '');

    if (!jwt || jwt.split('.').length !== 3) {
      throw new AuthCaptureError(
        'NO_TOKEN',
        'captured Authorization header did not contain a JWT',
      );
    }

    // 11. Decode claims.
    let claims: JwtClaims;
    try {
      claims = decodeJwt(jwt);
    } catch {
      throw new AuthCaptureError(
        'NO_TOKEN',
        'captured Bearer token could not be decoded as JWT',
      );
    }

    const expiresAt = new Date(claims.exp * 1000).toISOString();
    const audience = claims.aud;
    const scopes =
      typeof claims.scp === 'string'
        ? claims.scp.split(/\s+/).filter((s) => s.length > 0)
        : [];

    // 12. Account fields — prefer JWT claims, fall back to /me if needed.
    let puid = stringOrEmpty(claims.oid) || stringOrEmpty(claims.puid);
    let tenantId = stringOrEmpty(claims.tid);
    let upn =
      stringOrEmpty(claims.preferred_username) || stringOrEmpty(claims.upn);

    if (!puid || !tenantId || !upn) {
      const meCookies = await context.cookies();
      const fallback = await fetchMeFallback(jwt, meCookies);
      if (!puid) puid = fallback.puid;
      if (!tenantId) tenantId = fallback.tenantId;
      if (!upn) upn = fallback.upn;
    }

    if (!puid || !tenantId) {
      throw new AuthCaptureError(
        'NO_TOKEN',
        'captured JWT is missing required account identifiers (oid/tid) and /me fallback did not supply them',
      );
    }

    // 13. Collect and filter cookies.
    const allCookies = await context.cookies();
    const cookies: Cookie[] = allCookies
      .filter((c) => isRelevantCookieDomain(c.domain))
      .map(toSchemaCookie);

    // 14. Build the final result.
    const account = { upn, puid, tenantId };
    const anchorMailbox = `PUID:${account.puid}@${account.tenantId}`;

    // 14b. Optionally capture SharePoint session from the same context
    // before teardown. Failure here is non-fatal for the Outlook part —
    // we surface it as an error if it fails, since the caller asked for it.
    let sharepointSession: SharepointSession | undefined;
    if (opts.sharepointHost && opts.sharepointHost.length > 0) {
      sharepointSession = await captureSharepointFromContext(
        context,
        opts.sharepointHost,
        opts.loginTimeoutMs,
      );
    }

    return {
      bearer: { token: jwt, expiresAt, audience, scopes },
      cookies,
      account,
      anchorMailbox,
      sharepoint: sharepointSession,
    };
  } finally {
    // 15. Always close the context — even on error, even on cancellation.
    try {
      await context.close();
    } catch {
      // Context may already be closed by the user or by a Playwright teardown.
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function stringOrEmpty(v: unknown): string {
  return typeof v === 'string' ? v : '';
}

function isRelevantCookieDomain(domain: string): boolean {
  if (!domain) return false;
  const d = domain.toLowerCase();
  return COOKIE_DOMAIN_SUFFIXES.some((suffix) => {
    const s = suffix.toLowerCase();
    if (s.startsWith('.')) {
      // Leading dot means "this domain or any subdomain".
      return d === s.slice(1) || d.endsWith(s);
    }
    return d === s || d.endsWith('.' + s);
  });
}

/**
 * Convert a Playwright cookie into the persisted schema shape. Playwright's
 * Cookie type already uses the same field names; we map SameSite strictly to
 * the three literal values the schema demands.
 */
function toSchemaCookie(c: {
  name: string;
  value: string;
  domain: string;
  path: string;
  expires: number;
  httpOnly: boolean;
  secure: boolean;
  sameSite?: 'Strict' | 'Lax' | 'None' | undefined;
}): Cookie {
  const sameSite: 'Strict' | 'Lax' | 'None' =
    c.sameSite === 'Strict' || c.sameSite === 'Lax' || c.sameSite === 'None'
      ? c.sameSite
      : 'Lax';

  return {
    name: c.name,
    value: c.value,
    domain: c.domain,
    path: c.path,
    expires: c.expires,
    httpOnly: c.httpOnly,
    secure: c.secure,
    sameSite,
  };
}

/**
 * Last-resort account info resolver. Calls GET /api/v2.0/me with the captured
 * bearer and cookies. Kept minimal — no retries, no streaming — because the
 * primary JWT claims path already covers the vast majority of tenants.
 */
async function fetchMeFallback(
  jwt: string,
  cookies: ReadonlyArray<{ name: string; value: string; domain: string }>,
): Promise<{ upn: string; puid: string; tenantId: string }> {
  const cookieHeader = cookies
    .filter((c) => isRelevantCookieDomain(c.domain))
    .map((c) => `${c.name}=${c.value}`)
    .join('; ');

  const headers: Record<string, string> = {
    Authorization: `Bearer ${jwt}`,
    Accept: 'application/json',
  };
  if (cookieHeader.length > 0) {
    headers['Cookie'] = cookieHeader;
  }

  let resp: Response;
  try {
    resp = await fetch('https://outlook.office.com/api/v2.0/me', { headers });
  } catch (err) {
    throw new AuthCaptureError(
      'NO_TOKEN',
      `failed to call /me fallback for account resolution: ${(err as Error).message}`,
    );
  }

  if (!resp.ok) {
    throw new AuthCaptureError(
      'NO_TOKEN',
      `/me fallback returned HTTP ${resp.status}`,
    );
  }

  let body: unknown;
  try {
    body = await resp.json();
  } catch {
    throw new AuthCaptureError(
      'NO_TOKEN',
      '/me fallback returned non-JSON body',
    );
  }

  const me = (body ?? {}) as Record<string, unknown>;

  // OWA /me typically exposes EmailAddress / Id; tenant is rarely on /me so we
  // still depend on a JWT claim for tid in practice. We do our best.
  const upn =
    stringOrEmpty(me['EmailAddress']) ||
    stringOrEmpty(me['userPrincipalName']) ||
    stringOrEmpty(me['mail']);
  const puid =
    stringOrEmpty(me['Id']) ||
    stringOrEmpty(me['id']) ||
    stringOrEmpty(me['oid']);
  const tenantId =
    stringOrEmpty(me['TenantId']) ||
    stringOrEmpty(me['tenantId']) ||
    stringOrEmpty(me['tid']);

  return { upn, puid, tenantId };
}
