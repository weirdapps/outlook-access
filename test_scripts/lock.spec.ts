// test_scripts/lock.spec.ts
//
// Unit tests for src/auth/lock.ts — acquireLock / release / stale cleanup.

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { acquireLock } from '../src/auth/lock';

describe('acquireLock', () => {
  let tmpRoot: string;

  beforeEach(() => {
    tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-cli-test-'));
  });

  afterEach(() => {
    try {
      fs.rmSync(tmpRoot, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  it('creates the lock file with the current PID; release deletes it', async () => {
    const lockPath = path.join(tmpRoot, 'my.lock');
    const release = await acquireLock(lockPath);
    expect(fs.existsSync(lockPath)).toBe(true);

    const content = fs.readFileSync(lockPath, 'utf8').trim();
    expect(Number.parseInt(content, 10)).toBe(process.pid);

    await release();
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('release() is idempotent — second call is a no-op', async () => {
    const lockPath = path.join(tmpRoot, 'idem.lock');
    const release = await acquireLock(lockPath);
    await release();
    await expect(release()).resolves.toBeUndefined();
  });

  it('throws when another live instance holds the lock', async () => {
    const lockPath = path.join(tmpRoot, 'contended.lock');
    const releaseFirst = await acquireLock(lockPath);
    try {
      await expect(acquireLock(lockPath)).rejects.toThrowError(
        /another outlook-cli instance holds the lock/,
      );
    } finally {
      await releaseFirst();
    }
  });

  it('treats a stale PID lock as cleanable and acquires successfully', async () => {
    const lockPath = path.join(tmpRoot, 'stale.lock');
    // Write a PID that is almost certainly dead (very large, unlikely to exist).
    // We pick 999999; test is tolerant if the PID happens to exist by writing
    // 2^31-1 as a backup strategy would need process.kill to actually check.
    fs.writeFileSync(lockPath, '999999\n', { mode: 0o600 });

    // Sanity: ensure that PID is NOT actually alive. If it somehow is, skip the
    // assertion part about stale cleanup with a clear message.
    let isAlive = false;
    try {
      process.kill(999999, 0);
      isAlive = true;
    } catch {
      isAlive = false;
    }
    if (isAlive) {
      // Unlikely — but if the test host genuinely has PID 999999, skip the
      // stale-cleanup assertion (still valid: acquire would throw contention).
      return;
    }

    const release = await acquireLock(lockPath);
    // After acquisition the file should contain *our* pid.
    const newContent = fs.readFileSync(lockPath, 'utf8').trim();
    expect(Number.parseInt(newContent, 10)).toBe(process.pid);
    await release();
    expect(fs.existsSync(lockPath)).toBe(false);
  });

  it('treats an empty/unreadable-PID lock as stale and acquires', async () => {
    const lockPath = path.join(tmpRoot, 'emptypid.lock');
    // File with no PID content at all.
    fs.writeFileSync(lockPath, '\n', { mode: 0o600 });
    const release = await acquireLock(lockPath);
    const content = fs.readFileSync(lockPath, 'utf8').trim();
    expect(Number.parseInt(content, 10)).toBe(process.pid);
    await release();
  });
});
