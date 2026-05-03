// test_scripts/config.spec.ts
//
// Unit tests for src/config/config.ts — loadConfig resolution, precedence and
// validation of mandatory fields. Covers AC-CONFIG-MISSING.
//
// No live network, no browser. Pure logic.

import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { loadConfig, ENV, DEFAULTS } from '../src/config/config';
import { ConfigurationError } from '../src/config/errors';

const ALL_ENV_KEYS = [
  ENV.HTTP_TIMEOUT_MS,
  ENV.LOGIN_TIMEOUT_MS,
  ENV.CHROME_CHANNEL,
  ENV.SESSION_FILE,
  ENV.PROFILE_DIR,
  ENV.TZ,
  ENV.CAL_FROM,
  ENV.CAL_TO,
];

describe('loadConfig', () => {
  let savedEnv: Record<string, string | undefined>;

  beforeEach(() => {
    savedEnv = {};
    for (const k of ALL_ENV_KEYS) {
      savedEnv[k] = process.env[k];
      delete process.env[k];
    }
  });

  afterEach(() => {
    for (const k of ALL_ENV_KEYS) {
      if (savedEnv[k] === undefined) {
        delete process.env[k];
      } else {
        process.env[k] = savedEnv[k];
      }
    }
  });

  // Exception recorded in CLAUDE.md (2026-04-21): httpTimeoutMs,
  // loginTimeoutMs, chromeChannel now fall back to documented defaults
  // when neither flag nor env is set.
  it('falls back to DEFAULTS.HTTP_TIMEOUT_MS when httpTimeoutMs is unresolved', () => {
    const cfg = loadConfig({});
    expect(cfg.httpTimeoutMs).toBe(DEFAULTS.HTTP_TIMEOUT_MS);
    expect(cfg.httpTimeoutMs).toBe(30_000);
  });

  it('falls back to DEFAULTS.LOGIN_TIMEOUT_MS when loginTimeoutMs is unresolved', () => {
    const cfg = loadConfig({});
    expect(cfg.loginTimeoutMs).toBe(DEFAULTS.LOGIN_TIMEOUT_MS);
    expect(cfg.loginTimeoutMs).toBe(300_000);
  });

  it('falls back to DEFAULTS.CHROME_CHANNEL when chromeChannel is unresolved', () => {
    const cfg = loadConfig({});
    expect(cfg.chromeChannel).toBe(DEFAULTS.CHROME_CHANNEL);
    expect(cfg.chromeChannel).toBe('chrome');
  });

  it('still throws ConfigurationError when httpTimeoutMs is present but malformed in env', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = 'not-a-number';
    try {
      loadConfig({});
      throw new Error('expected throw');
    } catch (err) {
      expect(err).toBeInstanceOf(ConfigurationError);
      expect((err as ConfigurationError).missingSetting).toBe('httpTimeoutMs');
    }
  });

  it('still throws ConfigurationError when httpTimeoutMs flag is a non-positive integer', () => {
    try {
      loadConfig({ httpTimeoutMs: 0 });
      throw new Error('expected throw');
    } catch (err) {
      expect(err).toBeInstanceOf(ConfigurationError);
      expect((err as ConfigurationError).missingSetting).toBe('httpTimeoutMs');
    }
  });

  it('resolves from env vars when no CLI flag is given', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = '1500';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'msedge';

    const cfg = loadConfig({});
    expect(cfg.httpTimeoutMs).toBe(1500);
    expect(cfg.loginTimeoutMs).toBe(60_000);
    expect(cfg.chromeChannel).toBe('msedge');
  });

  it('CLI flag wins over env var (precedence)', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = '1500';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'msedge';

    const cfg = loadConfig({
      httpTimeoutMs: 9999,
      loginTimeoutMs: 12345,
      chromeChannel: 'chrome',
    });
    expect(cfg.httpTimeoutMs).toBe(9999);
    expect(cfg.loginTimeoutMs).toBe(12345);
    expect(cfg.chromeChannel).toBe('chrome');
  });

  it('rejects non-integer env value for httpTimeoutMs (NaN)', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = 'not-a-number';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    expect(() => loadConfig({})).toThrowError(ConfigurationError);
  });

  it('rejects negative integers for mandatory timeouts', () => {
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    expect(() => loadConfig({ httpTimeoutMs: -100 })).toThrowError(ConfigurationError);
  });

  it('rejects zero for mandatory timeouts', () => {
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    expect(() => loadConfig({ httpTimeoutMs: 0 })).toThrowError(ConfigurationError);
  });

  it('rejects NaN via CLI flag (not a finite integer)', () => {
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    expect(() => loadConfig({ httpTimeoutMs: Number.NaN })).toThrowError(ConfigurationError);
  });

  it('populates optional fields with defaults when nothing overrides them', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = '5000';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';

    const cfg = loadConfig({});
    expect(typeof cfg.sessionFilePath).toBe('string');
    expect(cfg.sessionFilePath.endsWith('/.outlook-cli/session.json')).toBe(true);
    expect(typeof cfg.profileDir).toBe('string');
    expect(cfg.profileDir.endsWith('/.outlook-cli/playwright-profile')).toBe(true);
    expect(typeof cfg.tz).toBe('string');
    expect(cfg.tz.length).toBeGreaterThan(0);
    expect(cfg.outputMode).toBe('json');
    expect(cfg.listMailTop).toBe(10);
    expect(cfg.listMailFolder).toBe('Inbox');
    expect(cfg.bodyMode).toBe('text');
    expect(cfg.calFrom).toBe('now');
    expect(cfg.calTo).toBe('now + 7d');
    expect(cfg.quiet).toBe(false);
    expect(cfg.noAutoReauth).toBe(false);
  });

  it('returns a frozen config object', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = '5000';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    const cfg = loadConfig({});
    expect(Object.isFrozen(cfg)).toBe(true);
  });

  it('rejects listMailTop outside 1..1000', () => {
    process.env[ENV.HTTP_TIMEOUT_MS] = '5000';
    process.env[ENV.LOGIN_TIMEOUT_MS] = '60000';
    process.env[ENV.CHROME_CHANNEL] = 'chrome';
    expect(() => loadConfig({ listMailTop: 0 })).toThrowError(ConfigurationError);
    expect(() => loadConfig({ listMailTop: 1001 })).toThrowError(ConfigurationError);
  });
});
