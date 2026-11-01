// src/http/sharepoint-client.ts
//
// HTTP client for SharePoint (and OneDrive-for-Business) hosts. Uses the
// Bearer token + cookies captured during `outlook-cli login --sharepoint-host`.
//
// Read-only: only GET is exposed for fetching shared file content. Errors
// are mapped to a typed SharepointHttpError so the caller can distinguish
// 401 (auth-required), 404/410 (stale), 403 (access-denied) from other
// HTTP failures.

export interface SharepointClientOpts {
  /** Optional: cookie-authenticated tenants have no Bearer. Sent only when present. */
  bearer?: string;
  cookies: string;
  timeoutMs: number;
  /**
   * The session's tenant host, e.g. "contoso.sharepoint.com". When set, only that
   * host and its OneDrive twin ("contoso-my.sharepoint.com") receive the session.
   * Any *.sharepoint.com host can be a stranger's tenant, so the suffix alone is
   * not enough once the tenant is known.
   */
  host?: string;
}

export interface SharepointBinaryResult {
  bytes: Buffer;
  contentType: string;
  size: number;
  /** Filename parsed from Content-Disposition; undefined if header missing. */
  filename?: string;
}

export class SharepointHttpError extends Error {
  constructor(
    public readonly status: number,
    public readonly url: string,
    msg: string,
  ) {
    super(msg);
    this.name = 'SharepointHttpError';
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

function parseContentDispositionFilename(header: string | null): string | undefined {
  if (!header) return undefined;
  // RFC 5987: filename*=UTF-8''encoded-name takes precedence
  const m1 = header.match(/filename\*=(?:UTF-8'')?([^;]+)/i);
  if (m1) {
    try {
      return decodeURIComponent(m1[1].trim().replace(/^"|"$/g, ''));
    } catch {
      /* fall through */
    }
  }
  // Fallback: filename="..."
  const m2 = header.match(/filename=("?)([^";]+)\1/i);
  if (m2) return m2[2].trim();
  return undefined;
}

/**
 * Throws unless `raw` is an https URL on a host allowed to receive the session.
 * The URL usually comes from a mail (a ReferenceAttachment SourceUrl), which the
 * sender controls, and the cookies authorize the whole tenant.
 */
export function assertSessionUrl(raw: string, tenantHost?: string): void {
  let u: URL;
  try {
    u = new URL(raw);
  } catch {
    throw new Error(`not a valid URL: "${raw}"`);
  }
  if (u.protocol !== 'https:') {
    throw new Error(`refusing non-https URL: "${raw}"`);
  }
  if (u.username || u.password) {
    throw new Error(`refusing URL with credentials in it: "${u.hostname}"`);
  }
  const h = u.hostname.toLowerCase();
  // Leading dot: "evilsharepoint.com" must not pass a suffix test.
  if (h !== 'sharepoint.com' && !h.endsWith('.sharepoint.com')) {
    throw new Error(
      `refusing to send the SharePoint session to non-SharePoint host "${u.hostname}"`,
    );
  }
  if (tenantHost) {
    const t = tenantHost.toLowerCase();
    const [label, ...rest] = t.split('.');
    const oneDrive = [`${label}-my`, ...rest].join('.');
    if (h !== t && h !== oneDrive) {
      throw new Error(`refusing to send the SharePoint session to another tenant "${u.hostname}"`);
    }
  }
}

export class SharepointClient {
  constructor(private readonly opts: SharepointClientOpts) {}

  async getBinary(absoluteUrl: string): Promise<SharepointBinaryResult> {
    assertSessionUrl(absoluteUrl, this.opts.host);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), this.opts.timeoutMs);
    try {
      const headers: Record<string, string> = {
        Accept: '*/*',
      };
      // Cookie-auth tenants have no Bearer; the FedAuth/rtFa cookies authorize
      // the request. Only send Authorization when a Bearer was captured.
      if (this.opts.bearer && this.opts.bearer.length > 0) {
        headers['Authorization'] = `Bearer ${this.opts.bearer}`;
      }
      if (this.opts.cookies && this.opts.cookies.length > 0) {
        headers['Cookie'] = this.opts.cookies;
      }
      const resp = await fetch(absoluteUrl, {
        method: 'GET',
        headers,
        signal: ctrl.signal,
        redirect: 'follow',
      });
      if (!resp.ok) {
        throw new SharepointHttpError(
          resp.status,
          absoluteUrl,
          `SharePoint GET ${absoluteUrl} → HTTP ${resp.status}`,
        );
      }
      const arrayBuf = await resp.arrayBuffer();
      return {
        bytes: Buffer.from(arrayBuf),
        contentType: resp.headers.get('content-type') ?? 'application/octet-stream',
        size: arrayBuf.byteLength,
        filename: parseContentDispositionFilename(resp.headers.get('content-disposition')),
      };
    } finally {
      clearTimeout(timer);
    }
  }
}
