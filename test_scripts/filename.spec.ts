// test_scripts/filename.spec.ts
//
// Unit tests for src/util/filename.ts — sanitizeAttachmentName,
// deduplicateFilename, assertWithinDir. Covers path-traversal defense used by
// download-attachments.

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import * as path from 'node:path';
import * as os from 'node:os';
import * as fs from 'node:fs';

import {
  sanitizeAttachmentName,
  deduplicateFilename,
  assertWithinDir,
  MAX_FILENAME_BYTES,
} from '../src/util/filename';

describe('sanitizeAttachmentName', () => {
  it('strips embedded path separators (forward slash)', () => {
    const s = sanitizeAttachmentName('foo/bar.txt');
    expect(s).not.toContain('/');
  });

  it('strips embedded path separators (backslash)', () => {
    const s = sanitizeAttachmentName('foo\\bar.txt');
    expect(s).not.toContain('\\');
  });

  it('removes traversal tokens like ".."', () => {
    const s = sanitizeAttachmentName('../etc/passwd');
    expect(s).not.toContain('..');
    expect(s).not.toContain('/');
    expect(s.length).toBeGreaterThan(0);
  });

  it('replaces Windows-illegal chars with underscore', () => {
    const s = sanitizeAttachmentName('a<b>c:d"e|f?g*h.txt');
    for (const ch of '<>:"|?*') {
      expect(s).not.toContain(ch);
    }
  });

  it('strips ASCII control chars', () => {
    const s = sanitizeAttachmentName('name\x00\x01\x1f.txt');
    // eslint-disable-next-line no-control-regex -- testing control char removal
    expect(s).not.toMatch(/[\x00-\x1F]/);
  });

  it('prefixes Windows reserved names with "_"', () => {
    expect(sanitizeAttachmentName('CON.txt')).toBe('_CON.txt');
    expect(sanitizeAttachmentName('con.TXT')).toBe('_con.TXT');
    expect(sanitizeAttachmentName('PRN')).toBe('_PRN');
    expect(sanitizeAttachmentName('COM1.log')).toBe('_COM1.log');
    expect(sanitizeAttachmentName('LPT9.data')).toBe('_LPT9.data');
  });

  it('truncates to ~MAX_FILENAME_BYTES preserving extension', () => {
    const base = 'a'.repeat(500);
    const name = `${base}.pdf`;
    const s = sanitizeAttachmentName(name);
    const byteLen = Buffer.byteLength(s, 'utf8');
    expect(byteLen).toBeLessThanOrEqual(MAX_FILENAME_BYTES);
    expect(s.endsWith('.pdf')).toBe(true);
  });

  it('returns "attachment" for empty input', () => {
    expect(sanitizeAttachmentName('')).toBe('attachment');
  });

  it('returns "attachment" for null/undefined coerced input', () => {
    expect(sanitizeAttachmentName(null as unknown as string)).toBe('attachment');
    expect(sanitizeAttachmentName(undefined as unknown as string)).toBe('attachment');
  });

  it('strips leading dots and trailing dots/spaces', () => {
    const s = sanitizeAttachmentName('...secret.txt...   ');
    expect(s.startsWith('.')).toBe(false);
    expect(s.endsWith('.')).toBe(false);
    expect(s.endsWith(' ')).toBe(false);
  });
});

describe('deduplicateFilename', () => {
  it('returns input unchanged when not present in the set', () => {
    const set = new Set<string>();
    expect(deduplicateFilename('foo.txt', set)).toBe('foo.txt');
  });

  it('adds " (1)" suffix before the extension on first collision', () => {
    const set = new Set<string>(['foo.txt']);
    expect(deduplicateFilename('foo.txt', set)).toBe('foo (1).txt');
  });

  it('increments the suffix sequentially', () => {
    const set = new Set<string>(['foo.txt', 'foo (1).txt']);
    expect(deduplicateFilename('foo.txt', set)).toBe('foo (2).txt');
  });

  it('handles names without extension', () => {
    const set = new Set<string>(['NOTES']);
    expect(deduplicateFilename('NOTES', set)).toBe('NOTES (1)');
  });
});

describe('assertWithinDir', () => {
  let baseDir: string;

  beforeEach(() => {
    baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-cli-test-'));
  });

  afterEach(() => {
    try {
      fs.rmSync(baseDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  it('returns absolute path inside baseDir for a normal name', () => {
    const resolved = assertWithinDir(baseDir, 'report.pdf');
    expect(path.isAbsolute(resolved)).toBe(true);
    expect(resolved.startsWith(path.resolve(baseDir))).toBe(true);
    expect(resolved.endsWith('report.pdf')).toBe(true);
  });

  it('throws "path traversal attempt" when filename escapes baseDir', () => {
    expect(() => assertWithinDir(baseDir, '../etc/passwd')).toThrowError(/path traversal attempt/);
    expect(() => assertWithinDir(baseDir, '..' + path.sep + 'escape.txt')).toThrowError(
      /path traversal attempt/,
    );
  });
});
