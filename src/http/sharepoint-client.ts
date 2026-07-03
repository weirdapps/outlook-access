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

export class SharepointClient {
  constructor(private readonly opts: SharepointClientOpts) {}

  async getBinary(absoluteUrl: string): Promise<SharepointBinaryResult> {
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
