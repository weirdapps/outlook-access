/**
 * HTTP-layer error classes.
 *
 * The hierarchy is flat by intent (easy to `instanceof`-discriminate at call
 * sites) and narrow: nothing here carries a cookie, a bearer token, or any
 * other secret. Every constructor runs its message through `redactString` as
 * a belt-and-suspenders guard against upstream-reflected secrets.
 *
 * Mapping to the CLI-wide error taxonomy (project-design.md §4) is performed
 * by the caller in `src/cli.ts`; the codes here are a superset of what the
 * cli wraps.
 */

import { OutlookCliError } from '../config/errors';
import { redactString } from '../util/redact';

// ---------------------------------------------------------------------------
// Base classes
// ---------------------------------------------------------------------------

/** Base for any error produced while performing an HTTP call to Outlook. */
export class OutlookHttpError extends Error {
  public readonly code: string;
  public readonly httpStatus: number;
  public readonly url: string;
  public readonly requestId?: string;

  constructor(init: {
    code: string;
    message: string;
    httpStatus: number;
    url: string;
    requestId?: string;
  }) {
    super(redactString(init.message));
    this.name = new.target.name;
    this.code = init.code;
    this.httpStatus = init.httpStatus;
    this.url = redactString(init.url);
    this.requestId = init.requestId;
  }
}

/**
 * 401 after the single automatic retry (or 401 with `--no-auto-reauth`).
 * Indicates that the stored bearer token is not acceptable to Outlook.
 *
 * The `reason` field discriminates the two paths so the caller can surface
 * the correct CLI-level error code (`AUTH_NO_REAUTH` vs `AUTH_401_AFTER_RETRY`).
 */
export type AuthErrorReason = 'NO_AUTO_REAUTH' | 'AFTER_RETRY';

export class AuthError extends OutlookHttpError {
  public readonly reason: AuthErrorReason;

  constructor(init: {
    message: string;
    url: string;
    httpStatus?: number;
    requestId?: string;
    reason: AuthErrorReason;
  }) {
    super({
      code: init.reason === 'NO_AUTO_REAUTH' ? 'AUTH_NO_REAUTH' : 'AUTH_REJECTED',
      message: init.message,
      httpStatus: init.httpStatus ?? 401,
      url: init.url,
      requestId: init.requestId,
    });
    this.reason = init.reason;
  }
}

/**
 * Non-401 4xx / 5xx from the upstream REST API. The `code` field is chosen by
 * the caller to match the status-specific error taxonomy:
 *
 *   404                       → 'NOT_FOUND'
 *   429                       → 'RATE_LIMITED'
 *   5xx                       → 'SERVER_ERROR'
 *   other 4xx (403, 409, ...) → 'API_ERROR' (or a status-specific string)
 *
 * Folder-feature discriminants (plan-002 P2, project-design §10.6):
 *
 *   'UPSTREAM_FOLDER_NOT_FOUND'   → 404 on `/me/MailFolders/{id}` during
 *                                    resolver path walk or id-kind lookup.
 *   'UPSTREAM_FOLDER_EXISTS'      → 400/409 with OData
 *                                    `error.code === 'ErrorFolderExists'`
 *                                    (see `isFolderExistsError`). The resolver
 *                                    reclassifies this into a `CollisionError`
 *                                    (non-idempotent) or a `PreExisting: true`
 *                                    recovery (idempotent).
 *   'UPSTREAM_PAGINATION_LIMIT'   → `listAll<T>` hit the 50-page safety cap
 *                                    or refused an off-host `@odata.nextLink`.
 */
export class ApiError extends OutlookHttpError {
  constructor(init: {
    code: string;
    message: string;
    httpStatus: number;
    url: string;
    requestId?: string;
  }) {
    super(init);
  }
}

// ---------------------------------------------------------------------------
// Network-layer failures (before a Response is received)
// ---------------------------------------------------------------------------

/**
 * Any failure that prevented a well-formed HTTP response from being received:
 *   - `fetch` threw a TypeError (DNS, TLS, connection reset, …)
 *   - `AbortError` from the per-request timeout signal
 *   - socket hang-up or ECONNRESET during body read
 *
 * Does NOT extend OutlookHttpError: there is no httpStatus.
 */
export class NetworkError extends Error {
  public readonly code = 'NETWORK';
  public readonly url: string;
  public readonly cause?: unknown;
  public readonly timedOut: boolean;

  constructor(init: {
    message: string;
    url: string;
    cause?: unknown;
    timedOut?: boolean;
  }) {
    super(redactString(init.message));
    this.name = 'NetworkError';
    this.url = redactString(init.url);
    this.cause = init.cause;
    this.timedOut = init.timedOut ?? false;
  }
}

// ---------------------------------------------------------------------------
// Status-to-code mapping helpers
// ---------------------------------------------------------------------------

/**
 * Map an HTTP status code (non-2xx, non-401) to the ApiError code string.
 *
 * 401 is intentionally absent: the HTTP client handles 401 specially (retry
 * once, then throw AuthError).
 */
export function codeForStatus(status: number): string {
  if (status === 403) return 'FORBIDDEN';
  if (status === 404) return 'NOT_FOUND';
  if (status === 409) return 'CONFLICT';
  if (status === 429) return 'RATE_LIMITED';
  if (status >= 500 && status <= 599) return 'SERVER_ERROR';
  if (status >= 400 && status <= 499) return 'API_ERROR';
  return 'UNEXPECTED_STATUS';
}

/**
 * Truncate a response-body snippet to at most `max` characters AND redact any
 * long base64-looking runs. Used when embedding upstream error bodies in
 * Error.message.
 */
export function truncateAndRedactBody(body: string, max = 512): string {
  if (!body) return '';
  const trimmed = body.length > max ? `${body.slice(0, max)}...` : body;
  return redactString(trimmed);
}

// ---------------------------------------------------------------------------
// CLI-layer folder errors
// ---------------------------------------------------------------------------

/**
 * Thrown when a folder-creation POST collides with an already-existing folder
 * (OData `error.code === 'ErrorFolderExists'`) AND the caller has NOT opted
 * into idempotent recovery (`--idempotent`). Shares exit code 6 with
 * `IoError` but keeps a distinct `instanceof` discriminator so
 * `cli.ts formatErrorJson` / `exitCodeFor` can emit a deterministic JSON
 * shape (`{code, path?, parentId?}`) without overloading the IO-error
 * vocabulary.
 *
 * Rationale: see ADR-13 (project-design §9) — the cause is NOT filesystem
 * IO (the existing exit-6 path is attachment-file collisions from
 * `download-attachments`). Keeping a dedicated class avoids code sprawl on
 * `IoError.code` and gives scripts a stable payload shape to match on.
 *
 * Canonical `code`: `'FOLDER_ALREADY_EXISTS'` — matches the
 * free-string discriminant convention used across the codebase
 * (e.g. `UpstreamError.code === 'UPSTREAM_HTTP_404'`,
 * `AuthError.code === 'AUTH_NO_REAUTH'`).
 *
 * Exit code 6.
 */
export class CollisionError extends OutlookCliError {
  public readonly code: string;
  public readonly exitCode: number = 6;
  /** Human-readable folder path segment or full path at the collision site. */
  public readonly path?: string;
  /** Id of the parent folder under which the collision occurred, if known. */
  public readonly parentId?: string;

  constructor(init: {
    code: string;
    message: string;
    path?: string;
    parentId?: string;
    cause?: unknown;
  }) {
    super(redactString(init.message), init.cause);
    this.name = 'CollisionError';
    this.code = init.code;
    this.path = init.path;
    this.parentId = init.parentId;
  }
}

// ---------------------------------------------------------------------------
// Folder-existence predicate
// ---------------------------------------------------------------------------

/**
 * Returns true when a parsed OData response body from
 * `POST /me/MailFolders` or `POST /me/MailFolders/{id}/childfolders`
 * indicates that a folder with the requested DisplayName already exists
 * under the target parent.
 *
 * The authoritative discriminator is `error.code === 'ErrorFolderExists'`.
 * Both HTTP 400 and HTTP 409 are observed across Exchange Online tenants
 * (see `docs/research/outlook-v2-folder-duplicate-error.md §4.1` and
 * project-design §10.6); the caller is responsible for gating the
 * status-code check. The message text is NOT used — it embeds the
 * folder name and is not locale-stable.
 *
 * The input is the parsed JSON body (the shape Outlook returns for any
 * error response on the v2.0 surface). Any non-object input, or a body
 * that does not match the `{ error: { code: "..." } }` shape, returns
 * `false`.
 */
export function isFolderExistsError(body: unknown): boolean {
  if (body === null || typeof body !== 'object') return false;
  const code: unknown = (body as { error?: { code?: unknown } })?.error?.code;
  return code === 'ErrorFolderExists';
}
