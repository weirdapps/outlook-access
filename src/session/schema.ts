// src/session/schema.ts
//
// Session file types and runtime validators.
// See docs/design/project-design.md §2.3 and refined-request-outlook-cli.md §7.2.

import { IoError } from '../config/errors';

/** Playwright cookie shape, persisted 1:1 in the session file. */
export interface Cookie {
  name: string;
  value: string;
  domain: string;
  path: string;
  /** Unix epoch seconds; -1 for session cookies. */
  expires: number;
  httpOnly: boolean;
  secure: boolean;
  sameSite: 'Strict' | 'Lax' | 'None';
}

export interface BearerInfo {
  /** Raw JWT. Never logged. */
  token: string;
  /** ISO8601 UTC, derived from JWT `exp` claim. */
  expiresAt: string;
  /** JWT `aud` claim. */
  audience: string;
  /** JWT `scp`/`scope` claim split on whitespace. Empty array if absent. */
  scopes: string[];
}

export interface Account {
  /** User principal name (e.g. "alice@contoso.com"). */
  upn: string;
  /** PUID/object ID from JWT `oid` (or `puid` if present). */
  puid: string;
  /** Tenant ID from JWT `tid`. */
  tenantId: string;
}

/** Matches refined spec §7.2 exactly. */
export interface SessionFile {
  /** Schema version. Currently 1. Bump on breaking changes. */
  version: 1;
  /** ISO8601 UTC, set at write time. */
  capturedAt: string;
  account: Account;
  bearer: BearerInfo;
  cookies: Cookie[];
  /** Pre-computed convenience: "PUID:<puid>@<tenantId>". */
  anchorMailbox: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal helpers — structural checks. Never log `bearer.token` or cookie
// `value` even on validation failure; messages reference only field names.
// ─────────────────────────────────────────────────────────────────────────────

function isNonEmptyString(v: unknown): v is string {
  return typeof v === 'string' && v.length > 0;
}

function isValidIsoDate(v: unknown): v is string {
  if (typeof v !== 'string' || v.length === 0) return false;
  const t = Date.parse(v);
  return Number.isFinite(t);
}

const JWT_SHAPE = /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/;

function isJwtShapedString(v: unknown): v is string {
  return typeof v === 'string' && JWT_SHAPE.test(v);
}

function isValidCookie(v: unknown): v is Cookie {
  if (v === null || typeof v !== 'object') return false;
  const c = v as Record<string, unknown>;
  if (typeof c.name !== 'string' || c.name.length === 0) return false;
  if (typeof c.value !== 'string') return false; // empty string is valid
  if (typeof c.domain !== 'string' || c.domain.length === 0) return false;
  if (typeof c.path !== 'string' || c.path.length === 0) return false;
  if (typeof c.expires !== 'number' || !Number.isFinite(c.expires)) return false;
  if (typeof c.httpOnly !== 'boolean') return false;
  if (typeof c.secure !== 'boolean') return false;
  if (c.sameSite !== 'Strict' && c.sameSite !== 'Lax' && c.sameSite !== 'None') {
    return false;
  }
  return true;
}

function isValidAccount(v: unknown): v is Account {
  if (v === null || typeof v !== 'object') return false;
  const a = v as Record<string, unknown>;
  return isNonEmptyString(a.upn) && isNonEmptyString(a.puid) && isNonEmptyString(a.tenantId);
}

function isValidBearer(v: unknown): v is BearerInfo {
  if (v === null || typeof v !== 'object') return false;
  const b = v as Record<string, unknown>;
  if (!isJwtShapedString(b.token)) return false;
  if (!isValidIsoDate(b.expiresAt)) return false;
  if (!isNonEmptyString(b.audience)) return false;
  if (!Array.isArray(b.scopes)) return false;
  for (const s of b.scopes) {
    if (typeof s !== 'string') return false;
  }
  return true;
}

/**
 * Non-aggressive runtime type guard for SessionFile. Returns true if the shape
 * is compatible with the documented schema. Use for permissive read-path
 * checks; use `validateSessionJson` when you want a hard error on mismatch.
 */
export function isValidSessionFile(x: unknown): x is SessionFile {
  if (x === null || typeof x !== 'object') return false;
  const s = x as Record<string, unknown>;
  if (s.version !== 1) return false;
  if (!isValidIsoDate(s.capturedAt)) return false;
  if (!isValidAccount(s.account)) return false;
  if (!isValidBearer(s.bearer)) return false;
  if (!Array.isArray(s.cookies)) return false;
  for (const c of s.cookies) {
    if (!isValidCookie(c)) return false;
  }
  if (typeof s.anchorMailbox !== 'string') return false;
  if (!s.anchorMailbox.startsWith('PUID:')) return false;
  if (!s.anchorMailbox.includes('@')) return false;
  return true;
}

/**
 * Strict validator: returns the narrowed SessionFile on success, throws
 * IoError("IO_SESSION_CORRUPT") on any structural mismatch. Used when callers
 * want to fail loudly (e.g. the CLI's write-then-read sanity path).
 *
 * Does NOT verify the JWT signature — only shape and types.
 */
export function validateSessionJson(raw: unknown): SessionFile {
  if (isValidSessionFile(raw)) {
    return raw;
  }
  throw new IoError({
    code: 'IO_SESSION_CORRUPT',
    message: 'Session file is corrupt or has unsupported schema.',
  });
}
