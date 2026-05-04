// src/config/config.ts
//
// Configuration resolver for the outlook-cli.
// See docs/design/project-design.md §2.1 and §5 for the normative contract.
//
// Hard rule (from project CLAUDE.md): never substitute a missing mandatory
// configuration value with a default or fallback. Throw ConfigurationError.
//
// Exception (recorded in CLAUDE.md "Project-specific exceptions to global
// rules", 2026-04-21): three runtime-plumbing settings — httpTimeoutMs,
// loginTimeoutMs, chromeChannel — now have defaults. Precedence is CLI
// flag > env var > default. No other setting is exempt.

import * as os from 'node:os';
import * as path from 'node:path';

import { ConfigurationError } from './errors';

export type OutputMode = 'json' | 'table';
export type BodyMode = 'html' | 'text' | 'none';

/**
 * The fully resolved configuration object for a single CLI invocation.
 * Every field marked "mandatory" in the refined spec §8 is non-optional here;
 * loadConfig() throws ConfigurationError if any such field is unresolved.
 */
export interface CliConfig {
  // Mandatory (throw ConfigurationError if unresolved)
  /** Per-REST-call HTTP timeout in milliseconds. */
  httpTimeoutMs: number;
  /** Max wall-clock time to wait for interactive login + first Bearer capture. */
  loginTimeoutMs: number;
  /** Playwright Chrome channel: e.g. "chrome", "msedge", "chrome-beta". */
  chromeChannel: string;

  // Optional with explicit defaults allowed by spec §8
  /** Path to the session file. Default: $HOME/.outlook-cli/session.json. */
  sessionFilePath: string;
  /** Path to the Playwright persistent profile directory (mode 0700). */
  profileDir: string;
  /** IANA timezone. Default: process.env.TZ ?? Intl system tz. */
  tz: string;
  /** Default output mode when neither --json nor --table is passed. */
  outputMode: OutputMode;
  /** Default --top for list-mail. Range 1..1000 (raised from 100 in v1.2.0). */
  listMailTop: number;
  /** Default --folder for list-mail. */
  listMailFolder: string;
  /** Default --body format for get-mail / get-event. */
  bodyMode: BodyMode;
  /** Calendar window start. ISO8601 or keyword. Default: "now". */
  calFrom: string;
  /** Calendar window end. ISO8601 or keyword. Default: "now + 7d". */
  calTo: string;
  /** When true, suppress progress messages on stderr. */
  quiet: boolean;
  /** When set, the session-file path override from --session-file flag. */
  sessionFileOverride?: string;
  /** When true, 401 or expired session does NOT trigger browser re-auth. */
  noAutoReauth: boolean;
  /** Optional path to a debug log file. When unset, no log file is written. */
  logFilePath?: string;
}

/** Partial of CliConfig representing flags collected from commander argv. */
export type CliFlags = Partial<{
  httpTimeoutMs: number;
  loginTimeoutMs: number;
  chromeChannel: string;
  sessionFilePath: string;
  profileDir: string;
  tz: string;
  outputMode: OutputMode;
  listMailTop: number;
  listMailFolder: string;
  bodyMode: BodyMode;
  calFrom: string;
  calTo: string;
  quiet: boolean;
  sessionFileOverride: string;
  noAutoReauth: boolean;
  logFilePath: string;
}>;

/**
 * Environment variable names consumed by loadConfig. Exported for tests and
 * for the configuration-guide document.
 */
export const ENV = {
  HTTP_TIMEOUT_MS: 'OUTLOOK_CLI_HTTP_TIMEOUT_MS',
  LOGIN_TIMEOUT_MS: 'OUTLOOK_CLI_LOGIN_TIMEOUT_MS',
  CHROME_CHANNEL: 'OUTLOOK_CLI_CHROME_CHANNEL',
  SESSION_FILE: 'OUTLOOK_CLI_SESSION_FILE',
  PROFILE_DIR: 'OUTLOOK_CLI_PROFILE_DIR',
  TZ: 'OUTLOOK_CLI_TZ',
  CAL_FROM: 'OUTLOOK_CLI_CAL_FROM',
  CAL_TO: 'OUTLOOK_CLI_CAL_TO',
} as const;

/**
 * Defaults for the three settings exempted from the no-fallback rule.
 * See CLAUDE.md "Project-specific exceptions to global rules" (2026-04-21).
 */
export const DEFAULTS = {
  HTTP_TIMEOUT_MS: 30_000,
  LOGIN_TIMEOUT_MS: 300_000,
  CHROME_CHANNEL: 'chrome',
} as const;

/**
 * Parse a numeric env var. Returns undefined if the env var is unset or empty.
 * Throws ConfigurationError if the value is present but not a valid integer.
 *
 * The settingName / sources arguments are used only to build a useful error.
 */
function parseIntEnv(
  envName: string,
  settingName: string,
  checkedSources: readonly string[],
): number | undefined {
  const raw = process.env[envName];
  if (raw === undefined || raw === '') {
    return undefined;
  }
  // Reject values containing any non-integer chars (leading/trailing whitespace
  // is tolerated; fractional / exponential / hex is not).
  const trimmed = raw.trim();
  if (!/^-?\d+$/.test(trimmed)) {
    throw new ConfigurationError(
      settingName,
      checkedSources,
      `env var ${envName} is not a valid integer (got ${JSON.stringify(raw)})`,
    );
  }
  const n = Number.parseInt(trimmed, 10);
  if (Number.isNaN(n)) {
    throw new ConfigurationError(
      settingName,
      checkedSources,
      `env var ${envName} is not a valid integer (got ${JSON.stringify(raw)})`,
    );
  }
  return n;
}

/**
 * Resolve an integer setting with a default. Precedence: CLI flag > env var
 * > default. Invalid-type values from flag/env still throw (the fallback
 * only covers the UNSET case, not the malformed case).
 *
 * Also validates positivity: a non-positive integer from flag/env is rejected.
 */
function resolveOptionalInt(
  settingName: string,
  flagValue: number | undefined,
  envName: string,
  flagLabel: string,
  defaultValue: number,
): number {
  const checkedSources: readonly string[] = [`${flagLabel} flag`, `${envName} env var`];
  let resolved: number | undefined;
  if (typeof flagValue === 'number') {
    if (!Number.isFinite(flagValue) || !Number.isInteger(flagValue)) {
      throw new ConfigurationError(
        settingName,
        checkedSources,
        `value from ${flagLabel} is not a finite integer (got ${String(flagValue)})`,
      );
    }
    resolved = flagValue;
  } else {
    resolved = parseIntEnv(envName, settingName, checkedSources);
  }
  if (resolved === undefined) {
    return defaultValue;
  }
  if (resolved <= 0) {
    throw new ConfigurationError(settingName, checkedSources, 'must be a positive integer');
  }
  return resolved;
}

/**
 * Resolve a string setting with a default. Precedence: CLI flag > env var
 * > default. Empty strings are treated as unset.
 */
function resolveOptionalString(
  flagValue: string | undefined,
  envName: string,
  defaultValue: string,
): string {
  const flag = typeof flagValue === 'string' && flagValue !== '' ? flagValue : undefined;
  const envRaw = process.env[envName];
  const env = typeof envRaw === 'string' && envRaw !== '' ? envRaw : undefined;
  return flag ?? env ?? defaultValue;
}

/**
 * Resolve the full CliConfig using precedence: CLI flag > environment variable
 * > explicit default (only where spec §8 allows one). Mandatory fields without
 * any resolved value throw ConfigurationError.
 */
export function loadConfig(cliFlags: CliFlags): CliConfig {
  // NOSONAR S3776 - config resolution with precedence rules
  // 1. Runtime-plumbing settings with defaults (see CLAUDE.md exception).
  //    Precedence: CLI flag > env var > default. Malformed flag/env still throws.
  const httpTimeoutMs = resolveOptionalInt(
    'httpTimeoutMs',
    cliFlags.httpTimeoutMs,
    ENV.HTTP_TIMEOUT_MS,
    '--timeout',
    DEFAULTS.HTTP_TIMEOUT_MS,
  );
  const loginTimeoutMs = resolveOptionalInt(
    'loginTimeoutMs',
    cliFlags.loginTimeoutMs,
    ENV.LOGIN_TIMEOUT_MS,
    '--login-timeout',
    DEFAULTS.LOGIN_TIMEOUT_MS,
  );
  const chromeChannel = resolveOptionalString(
    cliFlags.chromeChannel,
    ENV.CHROME_CHANNEL,
    DEFAULTS.CHROME_CHANNEL,
  );

  // 2. Optional with explicit defaults (spec §8 allows these)
  const home = os.homedir();

  const sessionFilePath =
    (typeof cliFlags.sessionFilePath === 'string' && cliFlags.sessionFilePath !== ''
      ? cliFlags.sessionFilePath
      : undefined) ??
    (process.env[ENV.SESSION_FILE] && process.env[ENV.SESSION_FILE] !== ''
      ? (process.env[ENV.SESSION_FILE] as string)
      : undefined) ??
    path.join(home, '.outlook-cli', 'session.json');

  const profileDir =
    (typeof cliFlags.profileDir === 'string' && cliFlags.profileDir !== ''
      ? cliFlags.profileDir
      : undefined) ??
    (process.env[ENV.PROFILE_DIR] && process.env[ENV.PROFILE_DIR] !== ''
      ? (process.env[ENV.PROFILE_DIR] as string)
      : undefined) ??
    path.join(home, '.outlook-cli', 'playwright-profile');

  const tz =
    (typeof cliFlags.tz === 'string' && cliFlags.tz !== '' ? cliFlags.tz : undefined) ??
    (process.env[ENV.TZ] && process.env[ENV.TZ] !== ''
      ? (process.env[ENV.TZ] as string)
      : undefined) ??
    Intl.DateTimeFormat().resolvedOptions().timeZone;

  const outputMode: OutputMode = cliFlags.outputMode ?? 'json';
  const listMailTop = cliFlags.listMailTop ?? 10;
  const listMailFolder = cliFlags.listMailFolder ?? 'Inbox';
  const bodyMode: BodyMode = cliFlags.bodyMode ?? 'text';

  const calFrom =
    (typeof cliFlags.calFrom === 'string' && cliFlags.calFrom !== ''
      ? cliFlags.calFrom
      : undefined) ??
    (process.env[ENV.CAL_FROM] && process.env[ENV.CAL_FROM] !== ''
      ? (process.env[ENV.CAL_FROM] as string)
      : undefined) ??
    'now';

  const calTo =
    (typeof cliFlags.calTo === 'string' && cliFlags.calTo !== '' ? cliFlags.calTo : undefined) ??
    (process.env[ENV.CAL_TO] && process.env[ENV.CAL_TO] !== ''
      ? (process.env[ENV.CAL_TO] as string)
      : undefined) ??
    'now + 7d';

  const quiet = cliFlags.quiet ?? false;
  const noAutoReauth = cliFlags.noAutoReauth ?? false;
  const sessionFileOverride =
    typeof cliFlags.sessionFileOverride === 'string' && cliFlags.sessionFileOverride !== ''
      ? cliFlags.sessionFileOverride
      : undefined;
  const logFilePath =
    typeof cliFlags.logFilePath === 'string' && cliFlags.logFilePath !== ''
      ? cliFlags.logFilePath
      : undefined;

  // 3. Sanity checks for optional numeric fields that still have bounds
  if (!Number.isInteger(listMailTop) || listMailTop < 1 || listMailTop > 1000) {
    throw new ConfigurationError(
      'listMailTop',
      ['--top flag'],
      'must be an integer between 1 and 1000',
    );
  }

  // 4. Build and freeze the result so downstream callers cannot mutate it.
  const cfg: CliConfig = {
    httpTimeoutMs,
    loginTimeoutMs,
    chromeChannel,
    sessionFilePath,
    profileDir,
    tz,
    outputMode,
    listMailTop,
    listMailFolder,
    bodyMode,
    calFrom,
    calTo,
    quiet,
    sessionFileOverride,
    noAutoReauth,
    logFilePath,
  };
  return Object.freeze(cfg);
}
