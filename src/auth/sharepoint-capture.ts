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

/**
 * Walks an existing context to a SharePoint host, captures the first
 * outbound Authorization: Bearer header, and returns a SharepointSession
 * ready to persist. Cookies for `host` (and its parent domain) are also
 * collected and serialized into the cookie header form.
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
    // Listen at the CONTEXT level, not the page level: modern SharePoint/Outlook
    // dispatch REST/Graph calls from a Service Worker, whose requests a
    // page-level listener (and in-page fetch/XHR hooks) can miss. Context-level
    // request events also see Service Worker traffic.
    const auth = await new Promise<string | null>((resolve) => {
      let done = false;
      const finish = (value: string | null): void => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        context.off('request', onRequest);
        resolve(value);
      };
      const onRequest = (req: Request): void => {
        try {
          if (!isSharepointBearerUrl(req.url())) return;
          const header = req.headers()['authorization'] ?? '';
          if (!/^Bearer\s+/i.test(header)) return;
          finish(header);
        } catch {
          /* best-effort — ignore malformed requests */
        }
      };
      const timer = setTimeout(() => finish(null), timeoutMs);
      context.on('request', onRequest);

      // Navigate to trigger SharePoint's authenticated REST calls. Don't await —
      // let the request listener race the navigation.
      page
        .goto(`https://${host}/_layouts/15/sharepoint.aspx`, {
          waitUntil: 'domcontentloaded',
          timeout: timeoutMs,
        })
        .catch(() => {
          // Navigation may itself error if SharePoint redirects oddly (e.g. via
          // the MCAS proxy); rely on the request listener instead.
        });
    });

    if (auth === null) {
      throw new SharepointCaptureError(
        'SHAREPOINT_TIMEOUT',
        `Timed out after ${timeoutMs}ms waiting for SharePoint Bearer header from ${host}`,
      );
    }

    const bearer = auth.replace(/^Bearer\s+/i, '');
    if (!bearer) {
      throw new SharepointCaptureError(
        'SHAREPOINT_NO_TOKEN',
        'Captured request had no Bearer token after stripping prefix',
      );
    }

    const claims = decodeJwt(bearer);
    const tokenExpiresAt = new Date(claims.exp * 1000).toISOString();

    // Collect cookies for the SharePoint host AND its parent domain
    // (e.g. *.sharepoint.com cookies are needed for cross-subdomain calls).
    const allCookies = await context.cookies();
    const parentDomain = host.split('.').slice(-2).join('.'); // sharepoint.com
    const sharepointCookies = allCookies
      .filter(
        (c) =>
          c.domain === host ||
          c.domain === `.${host}` ||
          c.domain === parentDomain ||
          c.domain === `.${parentDomain}`,
      )
      .map((c) => `${c.name}=${c.value}`)
      .join('; ');

    return {
      version: 1,
      host,
      bearer,
      cookies: sharepointCookies,
      capturedAt: new Date().toISOString(),
      tokenExpiresAt,
    };
  } finally {
    await page.close().catch(() => {
      /* tolerate teardown errors */
    });
  }
}
