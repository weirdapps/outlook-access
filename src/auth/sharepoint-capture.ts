// src/auth/sharepoint-capture.ts
//
// Capture a SharePoint Bearer + cookies from an existing Playwright
// BrowserContext (typically the same one that just captured the Outlook
// session). Uses page.waitForRequest to listen at the network layer rather
// than injecting an in-page hook — this is sufficient because SharePoint
// sends Authorization: Bearer on its REST/MSGraph calls during the initial
// page load.

import type { BrowserContext } from 'playwright';

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

  const page = await context.newPage();
  try {
    const requestPromise = page.waitForRequest(
      (req) => {
        try {
          const url = req.url();
          if (!url.startsWith(`https://${host}/`)) return false;
          const auth = req.headers()['authorization'] ?? '';
          return auth.startsWith('Bearer ');
        } catch {
          return false;
        }
      },
      { timeout: timeoutMs },
    );

    // Navigate. Don't await — let the request listener race.
    page
      .goto(`https://${host}/_layouts/15/sharepoint.aspx`, {
        waitUntil: 'domcontentloaded',
        timeout: timeoutMs,
      })
      .catch(() => {
        // Navigation may itself error if SharePoint redirects oddly; rely on
        // the request promise instead.
      });

    let request;
    try {
      request = await requestPromise;
    } catch {
      throw new SharepointCaptureError(
        'SHAREPOINT_TIMEOUT',
        `Timed out after ${timeoutMs}ms waiting for SharePoint Bearer header from ${host}`,
      );
    }

    const auth = request.headers()['authorization'] ?? '';
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
