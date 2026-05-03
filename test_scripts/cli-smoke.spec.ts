// test_scripts/cli-smoke.spec.ts
//
// Integration smoke-tests for the CLI entrypoint (src/cli.ts).
// Each test spawns the CLI as a child process via ts-node and asserts on
// stdout / stderr / exit code.

import { spawnSync, SpawnSyncReturns } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { describe, expect, it, vi } from 'vitest';

// CLI spawns are slower than unit tests; give each test plenty of headroom.
vi.setConfig({ testTimeout: 60_000 });

const PROJECT_ROOT = path.resolve(__dirname, '..');

interface RunOptions {
  env?: NodeJS.ProcessEnv;
  /** When true, clear all existing env; otherwise inherit process.env. */
  cleanEnv?: boolean;
}

function runCli(args: string[], options: RunOptions = {}): SpawnSyncReturns<string> {
  const baseEnv = options.cleanEnv ? { PATH: process.env.PATH ?? '' } : { ...process.env };
  return spawnSync('npx', ['ts-node', 'src/cli.ts', ...args], {
    env: { ...baseEnv, ...(options.env ?? {}) },
    cwd: PROJECT_ROOT,
    encoding: 'utf8',
    timeout: 30_000,
  });
}

function makeTempHome(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-cli-smoke-'));
}

/** Strip the env vars that would provide mandatory config, so CONFIG_MISSING
 *  fires. We don't pass `cleanEnv` because we need PATH + node to work. */
function clearedConfigEnv(): NodeJS.ProcessEnv {
  return {
    OUTLOOK_CLI_HTTP_TIMEOUT_MS: '',
    OUTLOOK_CLI_LOGIN_TIMEOUT_MS: '',
    OUTLOOK_CLI_CHROME_CHANNEL: '',
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('cli smoke tests', () => {
  it('(1) --help works without any env vars and lists all subcommands', () => {
    const r = runCli(['--help'], { env: clearedConfigEnv() });
    expect(r.status).toBe(0);
    const out = r.stdout;
    expect(out).toContain('list-mail');
    expect(out).toContain('get-mail');
    expect(out).toContain('download-attachments');
    expect(out).toContain('list-calendar');
    expect(out).toContain('get-event');
    expect(out).toContain('login');
    expect(out).toContain('auth-check');
    // Phase-7 folder-management subcommands must also be discoverable.
    expect(out).toContain('list-folders');
    expect(out).toContain('find-folder');
    expect(out).toContain('create-folder');
    expect(out).toContain('move-mail');
  });

  it('(1a) list-folders --help exposes the --parent, --recursive, --first-match flags', () => {
    const r = runCli(['list-folders', '--help'], { env: clearedConfigEnv() });
    expect(r.status).toBe(0);
    const out = r.stdout;
    expect(out).toContain('--parent');
    expect(out).toContain('--recursive');
    expect(out).toContain('--first-match');
  });

  it('(1b) move-mail --help exposes the --to and --continue-on-error flags', () => {
    const r = runCli(['move-mail', '--help'], { env: clearedConfigEnv() });
    expect(r.status).toBe(0);
    const out = r.stdout;
    expect(out).toContain('--to');
    expect(out).toContain('--continue-on-error');
  });

  it('(2) --version prints a semver and exits 0', () => {
    const r = runCli(['--version'], { env: clearedConfigEnv() });
    expect(r.status).toBe(0);
    expect(r.stdout.trim()).toMatch(/^\d+\.\d+\.\d+/);
  });

  it('(3) unknown subcommand exits 2 with USAGE_ERROR', () => {
    const r = runCli(['this-subcommand-does-not-exist'], {
      env: clearedConfigEnv(),
    });
    expect(r.status).toBe(2);
    const combined = `${r.stdout}\n${r.stderr}`;
    expect(combined).toMatch(/unknown command/i);
    expect(r.stderr).toContain('"USAGE_ERROR"');
  });

  it('(4) list-mail without config no longer exits 3 — defaults resolve and it falls through to AUTH (exit 4)', () => {
    // CLAUDE.md exception (2026-04-21): httpTimeoutMs / loginTimeoutMs /
    // chromeChannel now have defaults, so config resolution no longer fails
    // when all three are unset. With --no-auto-reauth and a fresh HOME the
    // next failure point is auth (exit 4).
    const home = makeTempHome();
    try {
      const r = runCli(['--no-auto-reauth', 'list-mail'], {
        env: {
          ...clearedConfigEnv(),
          HOME: home,
          OUTLOOK_CLI_SESSION_FILE: '',
          OUTLOOK_CLI_PROFILE_DIR: '',
        },
      });
      expect(r.status).toBe(4);
      expect(r.stderr).not.toContain('"CONFIG_MISSING"');
      expect(r.stderr).toContain('"code":');
    } finally {
      try {
        fs.rmSync(home, { recursive: true, force: true });
      } catch {
        /* best-effort cleanup */
      }
    }
  });

  it('(5) list-mail with config but no session + --no-auto-reauth → exit 4', () => {
    const home = makeTempHome();
    try {
      const r = runCli(['--no-auto-reauth', 'list-mail'], {
        env: {
          OUTLOOK_CLI_HTTP_TIMEOUT_MS: '30000',
          OUTLOOK_CLI_LOGIN_TIMEOUT_MS: '300000',
          OUTLOOK_CLI_CHROME_CHANNEL: 'chrome',
          // Override HOME so the CLI looks for the session file under a
          // fresh empty directory — guarantees cache miss.
          HOME: home,
          // Neutralize any override variables.
          OUTLOOK_CLI_SESSION_FILE: '',
          OUTLOOK_CLI_PROFILE_DIR: '',
        },
      });
      expect(r.status).toBe(4);
      const stderr = r.stderr;
      expect(stderr).toContain('"code":');
      // Spec says AUTH_NO_REAUTH; accept either code variant to stay robust.
      expect(stderr).toMatch(/AUTH_NO_REAUTH|AUTH_REJECTED|AUTH_401_AFTER_RETRY/);
    } finally {
      try {
        fs.rmSync(home, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  });
});
