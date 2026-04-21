/**
 * Redaction helpers. Used by every module that logs or embeds strings into
 * error messages. The goal is to make it structurally impossible for a bearer
 * token, cookie value, or other long opaque secret to leak into:
 *   - stderr messages
 *   - Error.message fields
 *   - log lines
 *
 * None of these helpers perform IO.
 */

const REDACTED = '[REDACTED]';

/**
 * Return a shallow copy of the headers object with sensitive values replaced
 * by "[REDACTED]". The lookup is case-insensitive.
 *
 * Redacted keys (case-insensitive):
 *   - authorization
 *   - cookie
 *   - set-cookie
 *   - x-ms-*-token           (wildcard: any header whose lowercase form starts
 *                             with "x-ms-" AND ends with "-token")
 *
 * x-anchormailbox is NOT considered secret and is preserved.
 *
 * The original object is not mutated. Keys are preserved with their original
 * casing so call-sites can still pretty-print them.
 */
export function redactHeaders(
  headers: Record<string, string>,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(headers)) {
    const lower = k.toLowerCase();
    if (isSensitiveHeaderName(lower)) {
      out[k] = REDACTED;
    } else {
      out[k] = v;
    }
  }
  return out;
}

function isSensitiveHeaderName(lowerName: string): boolean {
  if (lowerName === 'authorization') return true;
  if (lowerName === 'cookie') return true;
  if (lowerName === 'set-cookie') return true;
  if (lowerName.startsWith('x-ms-') && lowerName.endsWith('-token')) return true;
  return false;
}

/**
 * Redact a JWT (or any other long opaque token) for human-readable logs.
 *
 * Format: `<first 10 chars>...<last 5 chars>`.
 *
 * - If the input is falsy or shorter than 16 chars, returns "[REDACTED]" — a
 *   short token is as-good-as-full disclosure once prefix+suffix are shown.
 * - The prefix may include "Bearer " since callers sometimes pass the header
 *   value; we do not strip it on purpose (the caller decides).
 */
export function redactJwt(token: string): string {
  if (!token || token.length < 16) return REDACTED;
  return `${token.slice(0, 10)}...${token.slice(-5)}`;
}

/**
 * Replace any long (>100 char) run of base64-URL / base64 characters inside an
 * arbitrary string with "[REDACTED]". This is the last-line-of-defense helper
 * used when formatting error messages derived from upstream body text — the
 * upstream may echo back our token or a session cookie in rare cases.
 *
 * Character class matches base64 + base64url alphabets plus the JWT separator:
 *   [A-Za-z0-9+/=_\-.]
 * so JWTs with their 3 base64url segments are captured as a single run.
 */
export function redactString(s: string): string {
  if (!s) return s;
  const re = /[A-Za-z0-9+/=_\-.]{101,}/g;
  return s.replace(re, REDACTED);
}
