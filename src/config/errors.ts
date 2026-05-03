// src/config/errors.ts
//
// Typed error classes for the outlook-cli tool.
// See docs/design/project-design.md §2.2 and §4 for the normative contract.
//
// Redaction contract: no constructor here may place the bearer token, cookie
// values, or any header dictionary containing them into `.message`, `.stack`,
// or `.cause`. Callers that wrap upstream errors must redact before constructing.

/**
 * Base class for every CLI-thrown error. Carries a stable machine-friendly
 * `code` and the CLI process exit code to emit.
 */
export abstract class OutlookCliError extends Error {
  /** Stable, machine-friendly code (e.g. "CONFIG_MISSING"). */
  public abstract readonly code: string;
  /** CLI exit code (see design §4). */
  public abstract readonly exitCode: number;
  /** Underlying cause, if any. MUST NOT leak tokens or cookie values. */
  public readonly cause?: unknown;

  constructor(message: string, cause?: unknown) {
    super(message);
    this.name = this.constructor.name;
    this.cause = cause;
  }
}

/**
 * Thrown when a mandatory configuration setting cannot be resolved, or when a
 * resolved configuration setting fails a sanity check (e.g. non-positive
 * integer, out-of-range value).
 *
 * Exit code 3.
 */
export class ConfigurationError extends OutlookCliError {
  public readonly code: string = 'CONFIG_MISSING';
  public readonly exitCode: number = 3;
  /** Name of the unresolved or invalid mandatory setting (e.g. "httpTimeoutMs"). */
  public readonly missingSetting: string;
  /** Ordered list of sources checked (e.g. ["--timeout flag", "OUTLOOK_CLI_HTTP_TIMEOUT_MS env var"]). */
  public readonly checkedSources: readonly string[];

  constructor(missingSetting: string, checkedSources: readonly string[], detail?: string) {
    const sourcesText = checkedSources.length > 0 ? checkedSources.join(', ') : '(no sources)';
    const base = detail
      ? `Mandatory setting "${missingSetting}" ${detail}. Checked: ${sourcesText}.`
      : `Mandatory setting "${missingSetting}" was not provided. Checked: ${sourcesText}.`;
    super(base);
    this.name = 'ConfigurationError';
    this.missingSetting = missingSetting;
    this.checkedSources = Object.freeze([...checkedSources]);
  }
}

/**
 * Thrown on auth capture failure: user cancellation, login timeout, second 401,
 * or when --no-auto-reauth is set and a re-auth would have been needed.
 *
 * Exit code 4.
 */
export class AuthError extends OutlookCliError {
  public readonly code:
    | 'AUTH_LOGIN_CANCELLED'
    | 'AUTH_LOGIN_TIMEOUT'
    | 'AUTH_401_AFTER_RETRY'
    | 'AUTH_NO_REAUTH';
  public readonly exitCode: number = 4;

  constructor(
    code: 'AUTH_LOGIN_CANCELLED' | 'AUTH_LOGIN_TIMEOUT' | 'AUTH_401_AFTER_RETRY' | 'AUTH_NO_REAUTH',
    message: string,
    cause?: unknown,
  ) {
    super(message, cause);
    this.name = 'AuthError';
    this.code = code;
  }
}

/**
 * Thrown on any non-401 upstream HTTP error, network error, timeout, or abort.
 *
 * Exit code 5.
 */
export class UpstreamError extends OutlookCliError {
  /** e.g. "UPSTREAM_HTTP_403", "UPSTREAM_TIMEOUT", "UPSTREAM_NETWORK". */
  public readonly code: string;
  public readonly exitCode: number = 5;
  public readonly httpStatus?: number;
  public readonly requestId?: string;
  /** URL with any query-string tokens redacted. */
  public readonly url?: string;

  constructor(init: {
    code: string;
    message: string;
    httpStatus?: number;
    requestId?: string;
    url?: string;
    cause?: unknown;
  }) {
    super(init.message, init.cause);
    this.name = 'UpstreamError';
    this.code = init.code;
    this.httpStatus = init.httpStatus;
    this.requestId = init.requestId;
    this.url = init.url;
  }
}

/**
 * Thrown on file-system errors: session file read/write, output dir, attachment
 * overwrite guard, etc.
 *
 * Exit code 6.
 */
export class IoError extends OutlookCliError {
  /** e.g. "IO_WRITE_EEXIST", "IO_MKDIR_EACCES", "IO_SESSION_WRITE". */
  public readonly code: string;
  public readonly exitCode: number = 6;
  public readonly path?: string;

  constructor(init: { code: string; message: string; path?: string; cause?: unknown }) {
    super(init.message, init.cause);
    this.name = 'IoError';
    this.code = init.code;
    this.path = init.path;
  }
}
