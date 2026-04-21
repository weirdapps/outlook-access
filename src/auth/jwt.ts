// src/auth/jwt.ts
//
// Unit 3 — Manual JWT payload decoder (no signature verification).
// Design ref: docs/design/project-design.md §2.5

/**
 * Minimal subset of standard + Microsoft claims we care about.
 * Unknown keys are preserved via the index signature.
 */
export interface JwtClaims {
  /** Expiration time — Unix epoch seconds. */
  exp: number;
  /** Audience (e.g. "https://outlook.office.com/"). */
  aud: string;
  /** Object ID — Microsoft-issued user id. */
  oid?: string;
  /** Alternate user id some OWA tokens carry. */
  puid?: string;
  /** Tenant ID (directory id). */
  tid?: string;
  /** User principal name (enterprise). */
  upn?: string;
  /** Microsoft "preferred username" alternative to upn. */
  preferred_username?: string;
  /** Space-delimited scopes. */
  scp?: string;
  /** Array-of-strings alternate scope claim. */
  roles?: string[];
  /** Allow extra keys — JWTs carry many arbitrary claims. */
  [k: string]: unknown;
}

/**
 * Decode the payload section of a JWT without verifying the signature.
 *
 * Accepts both a raw JWT and one prefixed with "Bearer ". In the latter case
 * the prefix is stripped before decoding.
 *
 * @throws Error with message "invalid JWT" on any structural / parse failure.
 */
export function decodeJwt(token: string): JwtClaims {
  if (typeof token !== 'string' || token.length === 0) {
    throw new Error('invalid JWT');
  }

  let raw = token;
  if (raw.startsWith('Bearer ')) {
    raw = raw.slice('Bearer '.length);
  }
  raw = raw.trim();

  const segments = raw.split('.');
  if (segments.length !== 3) {
    throw new Error('invalid JWT');
  }

  const payloadSegment = segments[1];
  if (!payloadSegment || payloadSegment.length === 0) {
    throw new Error('invalid JWT');
  }

  // base64url → base64: replace URL-safe chars and pad.
  let b64 = payloadSegment.replace(/-/g, '+').replace(/_/g, '/');
  const pad = b64.length % 4;
  if (pad === 2) {
    b64 += '==';
  } else if (pad === 3) {
    b64 += '=';
  } else if (pad === 1) {
    // 1-char remainder is never valid base64.
    throw new Error('invalid JWT');
  }

  let json: string;
  try {
    json = Buffer.from(b64, 'base64').toString('utf8');
  } catch {
    throw new Error('invalid JWT');
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('invalid JWT');
  }

  if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('invalid JWT');
  }

  const claims = parsed as Record<string, unknown>;

  // Minimal sanity: exp is a number, aud is a string. Other fields are optional.
  if (typeof claims.exp !== 'number' || !Number.isFinite(claims.exp)) {
    throw new Error('invalid JWT');
  }
  if (typeof claims.aud !== 'string') {
    throw new Error('invalid JWT');
  }

  return claims as JwtClaims;
}
