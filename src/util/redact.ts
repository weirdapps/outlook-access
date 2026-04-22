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
 * arbitrary string with "[REDACTED]". Also redacts message body content (HTML
 * or plain) that may appear in echoed-back JSON error bodies — see
 * `redactMessageBodies` for the body patterns covered.
 *
 * Character class matches base64 + base64url alphabets plus the JWT separator:
 *   [A-Za-z0-9+/=_\-.]
 * so JWTs with their 3 base64url segments are captured as a single run.
 */
export function redactString(s: string): string {
  if (!s) return s;
  const tokenRe = /[A-Za-z0-9+/=_\-.]{101,}/g;
  return redactMessageBodies(s.replace(tokenRe, REDACTED));
}

const REDACTED_BODY = '[REDACTED-BODY]';

/**
 * Redact email message body content from a JSON-shaped string. Targets the
 * shapes M365 may echo back inside an error response body when our send-mail
 * payload trips a server-side validation rule:
 *
 *   "Body":{"ContentType":"HTML","Content":"<the html>"}   → Content value redacted
 *   "Body":{"ContentType":"Text","Content":"the text"}     → Content value redacted
 *   "HtmlBody":"<the html>"                                 → value redacted
 *   "TextBody":"the text"                                   → value redacted
 *
 * The patterns are intentionally permissive (allow embedded escaped quotes)
 * but bounded so the regex doesn't catastrophically backtrack on adversarial
 * input. ContentType / Subject / recipients are intentionally NOT redacted —
 * those are debugging metadata.
 */
export function redactMessageBodies(s: string): string {
  if (!s) return s;
  let out = s;
  // "Body":{ ... "Content":"..." ... }
  out = out.replace(
    /("Body"\s*:\s*\{[^}]*"Content"\s*:\s*")([^"\\]|\\.){0,20000}"/g,
    `$1${REDACTED_BODY}"`,
  );
  // "HtmlBody":"..." or "TextBody":"..."
  out = out.replace(
    /("(?:HtmlBody|TextBody)"\s*:\s*")([^"\\]|\\.){0,20000}"/g,
    `$1${REDACTED_BODY}"`,
  );
  return out;
}
