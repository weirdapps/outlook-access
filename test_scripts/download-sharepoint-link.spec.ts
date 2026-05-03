// test_scripts/download-sharepoint-link.spec.ts
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { describe, it, expect, vi } from 'vitest';
import { run, SharepointSessionMissingError } from '../src/commands/download-sharepoint-link';
import { saveSharepointSession, SharepointSession } from '../src/session/sharepoint-schema';
import { SharepointHttpError } from '../src/http/sharepoint-client';

function freshSession(host = 'nbg.sharepoint.com'): SharepointSession {
  return {
    version: 1,
    host,
    bearer: 'redacted',
    cookies: 'rtFa=a',
    capturedAt: new Date().toISOString(),
    tokenExpiresAt: new Date(Date.now() + 3600_000).toISOString(),
  };
}

function expiredSession(): SharepointSession {
  return {
    ...freshSession(),
    tokenExpiresAt: new Date(Date.now() - 1000).toISOString(),
  };
}

async function setupTempSession(session: SharepointSession): Promise<{
  outDir: string;
  sessionPath: string;
}> {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sp-dl-'));
  const sessionPath = path.join(tmp, 'sharepoint-session.json');
  await saveSharepointSession(sessionPath, session);
  const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sp-dl-out-'));
  return { outDir, sessionPath };
}

describe('download-sharepoint-link', () => {
  it('saves the file using filename from Content-Disposition', async () => {
    const { outDir, sessionPath } = await setupTempSession(freshSession());
    const fakeClient = {
      getBinary: vi.fn(async () => ({
        bytes: Buffer.from('hello world'),
        contentType: 'application/pdf',
        size: 11,
        filename: 'report.pdf',
      })),
    };
    const result = await run(
      {
        sharepointSessionPath: sessionPath,
        httpTimeoutMs: 30_000,
        createSharepointClient: () => fakeClient as never,
      },
      'https://nbg.sharepoint.com/sites/foo/Eabc',
      { out: outDir },
    );

    expect(result.saved.length).toBe(1);
    expect(result.saved[0].name).toBe('report.pdf');
    expect(fs.existsSync(path.join(outDir, 'report.pdf'))).toBe(true);
  });

  it('falls back to URL-derived filename when no Content-Disposition', async () => {
    const { outDir, sessionPath } = await setupTempSession(freshSession());
    const fakeClient = {
      getBinary: vi.fn(async () => ({
        bytes: Buffer.from('x'),
        contentType: 'application/pdf',
        size: 1,
      })),
    };
    const result = await run(
      {
        sharepointSessionPath: sessionPath,
        httpTimeoutMs: 30_000,
        createSharepointClient: () => fakeClient as never,
      },
      'https://nbg.sharepoint.com/sites/foo/Documents/Report.pdf',
      { out: outDir },
    );
    expect(result.saved[0].name).toBe('Report.pdf');
  });

  it('returns skipped record on 404 (no throw)', async () => {
    const { outDir, sessionPath } = await setupTempSession(freshSession());
    const fakeClient = {
      getBinary: vi.fn(async () => {
        throw new SharepointHttpError(404, 'https://x', 'not found');
      }),
    };
    const result = await run(
      {
        sharepointSessionPath: sessionPath,
        httpTimeoutMs: 30_000,
        createSharepointClient: () => fakeClient as never,
      },
      'https://nbg.sharepoint.com/missing',
      { out: outDir },
    );
    expect(result.saved.length).toBe(0);
    expect(result.skipped.length).toBe(1);
    expect(result.skipped[0].reason).toBe('not-found');
    expect(result.skipped[0].status).toBe(404);
  });

  it('returns skipped record on 403 (access-denied)', async () => {
    const { outDir, sessionPath } = await setupTempSession(freshSession());
    const fakeClient = {
      getBinary: vi.fn(async () => {
        throw new SharepointHttpError(403, 'https://x', 'forbidden');
      }),
    };
    const result = await run(
      {
        sharepointSessionPath: sessionPath,
        httpTimeoutMs: 30_000,
        createSharepointClient: () => fakeClient as never,
      },
      'https://nbg.sharepoint.com/locked',
      { out: outDir },
    );
    expect(result.skipped[0].reason).toBe('access-denied');
  });

  it('throws SharepointSessionMissingError when session file does not exist', async () => {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sp-dl-'));
    await expect(
      run(
        { sharepointSessionPath: '/tmp/does-not-exist-' + Date.now(), httpTimeoutMs: 1000 },
        'https://x.sharepoint.com/y',
        { out: outDir },
      ),
    ).rejects.toThrow(SharepointSessionMissingError);
  });

  it('throws SharepointSessionMissingError when session is expired', async () => {
    const { outDir, sessionPath } = await setupTempSession(expiredSession());
    let err: unknown;
    try {
      await run(
        { sharepointSessionPath: sessionPath, httpTimeoutMs: 1000 },
        'https://x.sharepoint.com/y',
        { out: outDir },
      );
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(SharepointSessionMissingError);
    expect((err as Error).message).toMatch(/expired/);
  });

  it('rejects when --out is missing', async () => {
    const { sessionPath } = await setupTempSession(freshSession());
    await expect(
      run(
        { sharepointSessionPath: sessionPath, httpTimeoutMs: 1000 },
        'https://x.sharepoint.com/y',
        {},
      ),
    ).rejects.toThrow();
  });
});
