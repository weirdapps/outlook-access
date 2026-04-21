// test_scripts/fs-atomic.spec.ts
//
// Unit tests for src/util/fs-atomic.ts — atomicWriteJson, readJsonFile.
// Covers AC-PERMS (file mode 0600) and AC-OVERWRITE-GUARD (overwrite=false).

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { atomicWriteJson, readJsonFile } from '../src/util/fs-atomic';
import { IoError } from '../src/config/errors';

describe('atomicWriteJson', () => {
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

  it('writes the expected JSON content (pretty + trailing newline)', async () => {
    const p = path.join(tmpRoot, 'out.json');
    await atomicWriteJson(p, { a: 1, b: 'two' });
    const raw = fs.readFileSync(p, 'utf8');
    expect(raw).toBe(JSON.stringify({ a: 1, b: 'two' }, null, 2) + '\n');
    expect(JSON.parse(raw)).toEqual({ a: 1, b: 'two' });
  });

  it('creates file with mode 0o600 by default', async () => {
    const p = path.join(tmpRoot, 'secret.json');
    await atomicWriteJson(p, { hello: 'world' });
    const s = fs.statSync(p);
    if (process.platform !== 'win32') {
      expect(s.mode & 0o777).toBe(0o600);
    }
  });

  it('honours an explicit mode option', async () => {
    const p = path.join(tmpRoot, 'shared.json');
    await atomicWriteJson(p, { x: 1 }, { mode: 0o640 });
    const s = fs.statSync(p);
    if (process.platform !== 'win32') {
      expect(s.mode & 0o777).toBe(0o640);
    }
  });

  it('overwrite=false throws IoError(IO_WRITE_EEXIST) when target exists', async () => {
    const p = path.join(tmpRoot, 'exists.json');
    await atomicWriteJson(p, { first: true });

    let caught: unknown;
    try {
      await atomicWriteJson(p, { second: true }, { overwrite: false });
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(IoError);
    expect((caught as IoError).code).toBe('IO_WRITE_EEXIST');

    // Original file must be untouched.
    const raw = fs.readFileSync(p, 'utf8');
    expect(JSON.parse(raw)).toEqual({ first: true });
  });

  it('overwrite=true (default) replaces an existing file', async () => {
    const p = path.join(tmpRoot, 'replace.json');
    await atomicWriteJson(p, { first: true });
    await atomicWriteJson(p, { second: true });
    expect(JSON.parse(fs.readFileSync(p, 'utf8'))).toEqual({ second: true });
  });
});

describe('readJsonFile', () => {
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

  it('returns null on ENOENT', async () => {
    const result = await readJsonFile(path.join(tmpRoot, 'missing.json'));
    expect(result).toBeNull();
  });

  it('throws IoError(IO_SESSION_CORRUPT) on invalid JSON', async () => {
    const p = path.join(tmpRoot, 'broken.json');
    fs.writeFileSync(p, '{ not json', 'utf8');
    let caught: unknown;
    try {
      await readJsonFile(p);
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(IoError);
    expect((caught as IoError).code).toBe('IO_SESSION_CORRUPT');
  });

  it('returns the parsed object on valid JSON', async () => {
    const p = path.join(tmpRoot, 'ok.json');
    fs.writeFileSync(p, JSON.stringify({ a: 1 }), 'utf8');
    const result = await readJsonFile<{ a: number }>(p);
    expect(result).toEqual({ a: 1 });
  });
});
