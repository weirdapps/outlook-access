// src/auth/sharepoint-capture.ts
//
// Capture a SharePoint Bearer + cookies from an existing Playwright
// BrowserContext (typically the same one that just captured the Outlook
// session). Listens at the CONTEXT level for the first Authorization: Bearer
// request to the SharePoint host — including MCAS-proxied (*.mcas.ms) and
// Service-Worker-dispatched requests, which a page-level or in-page hook
// misses. SharePoint emits the Bearer on its REST/MSGraph calls during the
// initial page load.

import type { BrowserContext, Request } from 'playwright';

import { decodeJwt } from './jwt';
import type { SharepointSession } from '../session/sharepoint-schema';

export type SharepointCaptureErrorCode =
  | 'SHAREPOINT_TIMEOUT'
  | 'SHAREPOINT_NO_TOKEN'
  | 'SHAREPOINT_INVALID_HOST';

export class SharepointCaptureError extends Error {
  public readonly code: SharepointCaptureErrorCode;

  constructor(code: SharepointCaptureErrorCode, message: string) {
    super(message);
    this.name = 'SharepointCaptureError';
    this.code = code;
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

const VALID_HOST_RE = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;

function validateHost(host: string): void {
  if (!VALID_HOST_RE.test(host) || !host.includes('sharepoint.com')) {
    throw new SharepointCaptureError(
      'SHAREPOINT_INVALID_HOST',
      `Invalid SharePoint host "${host}" — expected something like "tenant.sharepoint.com"`,
    );
  }
}

// Cookie-auth sessions have no JWT to expire against. FedAuth/rtFa cookies carry
// their own expiry; when they're session cookies (expires = -1) we fall back to a
// conservative window. auth-renew re-captures every ~15 min, so this is only a
// backstop for the expiry pre-check and the health-check alert.
const COOKIE_FALLBACK_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

interface ExpiringCookie {
  name: string;
  expires?: number; // Unix seconds; -1 or absent for session cookies
}

/** Derive the session expiry: JWT exp when a Bearer was captured, else the
 * FedAuth (then rtFa) cookie expiry, else a conservative fallback window. */
function deriveTokenExpiry(bearer: string | undefined, cookies: ExpiringCookie[]): string {
  if (bearer) {
    try {
      return new Date(decodeJwt(bearer).exp * 1000).toISOString();
    } catch {
      /* malformed bearer — fall through to cookie-based expiry */
    }
  }
  for (const name of ['FedAuth', 'rtFa']) {
    const c = cookies.find((k) => k.name.toLowerCase() === name.toLowerCase());
    if (c && typeof c.expires === 'number' && c.expires > 0) {
      return new Date(c.expires * 1000).toISOString();
    }
  }
  return new Date(Date.now() + COOKIE_FALLBACK_TTL_MS).toISOString();
}

/**
 * Walks an existing context to a SharePoint host and returns a SharepointSession
 * ready to persist. Cookies for `host` (and its parent domain) are collected and
 * serialized into the cookie header form; a Bearer is captured best-effort.
 *
 * Cookie-auth tenants (e.g. MCAS-gated) never emit a SharePoint Bearer, so a
 * missing Bearer is NOT an error — the FedAuth/rtFa cookies authorize downloads.
 * Only a total absence of auth (no Bearer AND no cookies) fails.
 *
 * Should be called AFTER the Outlook session is captured — by then the
 * persistent context already has Microsoft sign-in cookies, so SharePoint
 * SSO completes silently.
 */
export async function captureSharepointFromContext(
  context: BrowserContext,
  host: string,
  timeoutMs: number,
): Promise<SharepointSession> {
  validateHost(host);

  // Match the SharePoint Bearer request whether it goes directly to the host or
  // is rewritten by MCAS (Microsoft Defender for Cloud Apps Conditional Access
  // App Control), which proxies requests through a "<original-fqdn>.mcas.ms"
  // domain — so the original host no longer prefixes the URL and a plain
  // host-prefix filter never matches (the headless/VPS timeout symptom).
  const tenant = host.split('.')[0].toLowerCase();
  const isSharepointBearerUrl = (url: string): boolean => {
    if (url.startsWith(`https://${host}/`)) return true;
    if (/\.mcas\.ms\//i.test(url)) {
      const lower = url.toLowerCase();
      return lower.includes(tenant) && lower.includes('sharepoint');
    }
    return false;
  };

  const page = await context.newPage();
  try {
    // Best-effort Bearer capture. Listen at the CONTEXT level (not page level):
    // modern SharePoint dispatches REST/Graph calls from a Service Worker, whose
    // requests a page-level listener can miss. We record the first Bearer seen
    // during navigation but never block on it — cookie-auth tenants emit none.
    let capturedAuth: string | null = null;
    const onRequest = (req: Request): void => {
      if (capturedAuth) return;
      try {
        if (!isSharepointBearerUrl(req.url())) return;
        const header = req.headers()['authorization'] ?? '';
        if (/^Bearer\s+/i.test(header)) capturedAuth = header;
      } catch {
        /* best-effort — ignore malformed requests */
      }
    };
    context.on('request', onRequest);
    try {
      // Await navigation (bounded by timeoutMs) so cookies are set and any Bearer
      // emitted during load is seen — then stop listening. This no longer hangs
      // for the full timeout on cookie-auth tenants (the old symptom).
      await page.goto(`https://${host}/_layouts/15/sharepoint.aspx`, {
        waitUntil: 'domcontentloaded',
        timeout: timeoutMs,
      });
    } catch {
      // Navigation may error under MCAS redirects; cookies are still set on the
      // context, so fall through and collect them.
    } finally {
      context.off('request', onRequest);
    }

    const bearer = capturedAuth ? (capturedAuth as string).replace(/^Bearer\s+/i, '') : undefined;

    // Collect cookies for the SharePoint host AND its parent domain
    // (e.g. *.sharepoint.com cookies are needed for cross-subdomain calls).
    const allCookies = await context.cookies();
    const parentDomain = host.split('.').slice(-2).join('.'); // sharepoint.com
    const relevant = allCookies.filter(
      (c) =>
        c.domain === host ||
        c.domain === `.${host}` ||
        c.domain === parentDomain ||
        c.domain === `.${parentDomain}`,
    );
    const cookies = relevant.map((c) => `${c.name}=${c.value}`).join('; ');

    if (!bearer && !cookies) {
      // No Bearer AND no cookies means sign-in established no SharePoint session
      // at all — surface it rather than persist a dead file.
      throw new SharepointCaptureError(
        'SHAREPOINT_NO_TOKEN',
        `No SharePoint auth captured for ${host} — sign-in may have failed (no Bearer, no cookies)`,
      );
    }

    return {
      version: 1,
      host,
      bearer,
      cookies,
      capturedAt: new Date().toISOString(),
      tokenExpiresAt: deriveTokenExpiry(bearer, relevant),
    };
  } finally {
    await page.close().catch(() => {
      /* tolerate teardown errors */
    });
  }
}
