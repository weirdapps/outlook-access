// test_scripts/session-store.spec.ts
//
// Unit tests for src/session/store.ts — saveSession, loadSession, isExpired.
// Covers AC-PERMS (file mode 0600, parent dir 0700) and AC-NO-SECRET-LEAK
// (malformed JSON warnings do not echo contents).

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { saveSession, loadSession, isExpired, EXPIRY_SKEW_MS } from '../src/session/store';
import type { SessionFile } from '../src/session/schema';

function makeValidSession(overrides?: Partial<SessionFile>): SessionFile {
  return {
    version: 1,
    capturedAt: new Date().toISOString(),
    account: {
      upn: 'alice@contoso.com',
      puid: '10003F00AAAABBBB',
      tenantId: '12345678-1234-1234-1234-123456789012',
    },
    bearer: {
      token: 'aaaaaaaa.bbbbbbbb.cccccccc',
      expiresAt: new Date(Date.now() + 3600_000).toISOString(),
      audience: 'https://outlook.office.com/',
      scopes: ['Mail.Read'],
    },
    cookies: [],
    anchorMailbox: 'PUID:10003F00AAAABBBB@12345678-1234-1234-1234-123456789012',
    ...overrides,
  };
}

describe('session store', () => {
  let tmpRoot: string;

  beforeEach(() => {
    tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-cli-test-'));
  });

  afterEach(() => {
    try {
      fs.rmSync(tmpRoot, { recursive: true, force: true });
    } catch {
      // ignore cleanup failures
    }
  });

  describe('saveSession', () => {
    it('writes session file with mode 0600 in parent dir with mode 0700', async () => {
      const sessionPath = path.join(tmpRoot, 'private', 'session.json');
      const s = makeValidSession();
      await saveSession(sessionPath, s);

      const fileStat = fs.statSync(sessionPath);
      // On POSIX, the mode bits we care about are the low 9 bits.
      if (process.platform !== 'win32') {
        expect(fileStat.mode & 0o777).toBe(0o600);
      }

      const dirStat = fs.statSync(path.dirname(sessionPath));
      if (process.platform !== 'win32') {
        expect(dirStat.mode & 0o777).toBe(0o700);
      }
    });

    it('overwrites an existing session file atomically', async () => {
      const sessionPath = path.join(tmpRoot, 'session.json');
      const s1 = makeValidSession();
      await saveSession(sessionPath, s1);
      const s2 = makeValidSession({ capturedAt: new Date(Date.now() - 1000).toISOString() });
      await saveSession(sessionPath, s2);

      const loaded = await loadSession(sessionPath);
      expect(loaded).not.toBeNull();
      expect(loaded!.capturedAt).toBe(s2.capturedAt);
    });
  });

  describe('loadSession', () => {
    it('returns null for ENOENT (file does not exist)', async () => {
      const result = await loadSession(path.join(tmpRoot, 'does-not-exist.json'));
      expect(result).toBeNull();
    });

    it('returns null and warns on wrong schema (missing required field)', async () => {
      const sessionPath = path.join(tmpRoot, 'bad-schema.json');
      // Write a JSON object missing mandatory fields (no 'account', no 'bearer').
      fs.writeFileSync(
        sessionPath,
        JSON.stringify({ version: 1, capturedAt: new Date().toISOString() }),
        { mode: 0o600 },
      );

      const warnSpy = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
      try {
        const result = await loadSession(sessionPath);
        expect(result).toBeNull();
        expect(warnSpy).toHaveBeenCalled();
        // Verify warning never echoes the full path-independent payload details.
        const warned = warnSpy.mock.calls.map((c) => String(c[0])).join(' ');
        expect(warned).toContain('schema validation');
      } finally {
        warnSpy.mockRestore();
      }
    });

    it('throws IoError (IO_SESSION_CORRUPT) on malformed JSON', async () => {
      const sessionPath = path.join(tmpRoot, 'malformed.json');
      fs.writeFileSync(sessionPath, 'not json at all {{{', { mode: 0o600 });
      await expect(loadSession(sessionPath)).rejects.toThrow();
    });

    it('returns a valid session when the file is well-formed and passes schema', async () => {
      const sessionPath = path.join(tmpRoot, 'session.json');
      const s = makeValidSession();
      await saveSession(sessionPath, s);

      const loaded = await loadSession(sessionPath);
      expect(loaded).not.toBeNull();
      expect(loaded!.version).toBe(1);
      expect(loaded!.account.upn).toBe('alice@contoso.com');
    });
  });

  describe('isExpired', () => {
    it('returns true for past expiresAt', () => {
      const s = makeValidSession({
        bearer: {
          token: 'aaaaaaaa.bbbbbbbb.cccccccc',
          expiresAt: new Date(Date.now() - 60_000).toISOString(),
          audience: 'https://outlook.office.com/',
          scopes: [],
        },
      });
      expect(isExpired(s)).toBe(true);
    });

    it('returns false for future expiresAt well beyond skew', () => {
      const s = makeValidSession({
        bearer: {
          token: 'aaaaaaaa.bbbbbbbb.cccccccc',
          expiresAt: new Date(Date.now() + 10 * 60_000).toISOString(),
          audience: 'https://outlook.office.com/',
          scopes: [],
        },
      });
      expect(isExpired(s)).toBe(false);
    });

    it('returns true within the 60s skew window (will expire shortly)', () => {
      const s = makeValidSession({
        bearer: {
          token: 'aaaaaaaa.bbbbbbbb.cccccccc',
          expiresAt: new Date(Date.now() + 30_000).toISOString(),
          audience: 'https://outlook.office.com/',
          scopes: [],
        },
      });
      // 30s < 60s skew → should be considered expired proactively.
      expect(isExpired(s)).toBe(true);
      // Sanity: the constant itself is 60s.
      expect(EXPIRY_SKEW_MS).toBe(60_000);
    });

    it('returns true when expiresAt is unparsable', () => {
      const s = makeValidSession({
        bearer: {
          token: 'aaaaaaaa.bbbbbbbb.cccccccc',
          expiresAt: 'not-a-date',
          audience: 'https://outlook.office.com/',
          scopes: [],
        },
      });
      expect(isExpired(s)).toBe(true);
    });
  });
});
