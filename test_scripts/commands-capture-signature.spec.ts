// test_scripts/commands-capture-signature.spec.ts

import { describe, it, expect, vi } from 'vitest';

import {
  run,
  extractSignature,
  UsageError,
} from '../src/commands/capture-signature';
import type { OutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';
import type { CliConfig } from '../src/config/config';

const MINIMAL_CONFIG = {
  httpTimeoutMs: 30000,
  loginTimeoutMs: 300000,
  chromeChannel: 'chrome',
  sessionFilePath: '/tmp/session.json',
  profileDir: '/tmp/profile',
  tz: 'UTC',
  outputMode: 'json',
  listMailTop: 10,
  listMailFolder: 'Inbox',
  bodyMode: 'text',
  calFrom: 'now',
  calTo: 'now + 7d',
  quiet: true,
  noAutoReauth: true,
} as unknown as CliConfig;

const SESSION: SessionFile = {
  version: 1,
  capturedAt: '2026-04-21T12:00:00.000Z',
  account: { upn: 'me@nbg.gr', puid: 'p', tenantId: 't' },
  bearer: {
    token: 'x.y.z',
    expiresAt: '2099-04-21T12:00:00.000Z',
    audience: 'https://outlook.office.com',
    scopes: ['Mail.Read'],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

function makeDeps(latestId: string | null, bodyHtml: string) {
  const client = {
    listMessagesInFolder: vi.fn(async () =>
      latestId ? [{ Id: latestId, Subject: 'last sent', SentDateTime: '2026-04-21T10:00:00Z' }] : [],
    ),
    getMessage: vi.fn(async () => ({
      Id: latestId ?? 'X',
      Subject: 'last sent',
      Body: { ContentType: 'HTML' as const, Content: bodyHtml },
    })),
  } as unknown as OutlookClient;

  const writeFile = vi.fn(async () => undefined);

  return {
    deps: {
      config: MINIMAL_CONFIG,
      sessionPath: '/tmp/session.json',
      loadSession: vi.fn(async () => SESSION),
      saveSession: vi.fn(async () => {}),
      doAuthCapture: vi.fn(async () => SESSION),
      createClient: vi.fn(() => client),
      writeFile,
      homeDir: () => '/tmp/fake-home',
    },
    client,
    writeFile,
  };
}

describe('extractSignature — heuristic priority', () => {
  it('1. <div id="Signature"> wins', () => {
    const html = `
      <p>Hi team,</p>
      <p>body content here</p>
      <div id="Signature">
        <b>Dimitrios Plessas</b><br>AGM, NBG
      </div>
    `;
    const out = extractSignature(html);
    expect(out.heuristic).toBe('div-signature');
    expect(out.signature).toContain('Dimitrios Plessas');
    expect(out.signature).not.toContain('body content here');
  });

  it('2. <div class="elementToProof"> when no Signature div', () => {
    const html = `
      <p>Some body.</p>
      <div class="elementToProof">
        <b>D. Plessas</b>
      </div>
    `;
    const out = extractSignature(html);
    expect(out.heuristic).toBe('div-elementtoproof');
    expect(out.signature).toContain('D. Plessas');
  });

  it('3. last <hr> when no signature div', () => {
    const html = `
      <p>Body content here.</p>
      <hr>
      <b>Dimitrios Plessas</b>
    `;
    const out = extractSignature(html);
    expect(out.heuristic).toBe('last-hr');
    expect(out.signature).toContain('Dimitrios Plessas');
    expect(out.signature).not.toContain('Body content here');
  });

  it('4. reply marker fallback — keeps content before "On ... wrote:"', () => {
    const html = `
      <p>Thanks.</p>
      <p>D.P.</p>
      <p>On Mon Apr 21 2026 alice@x.com wrote:</p>
      <blockquote>previous email body</blockquote>
    `;
    const out = extractSignature(html);
    expect(out.heuristic).toBe('reply-marker');
    expect(out.signature).toContain('D.P.');
    expect(out.signature).not.toContain('previous email body');
  });

  it('5. whole body fallback when nothing matches', () => {
    const html = '<p>just one paragraph no markers</p>';
    const out = extractSignature(html);
    expect(out.heuristic).toBe('whole-body');
    expect(out.signature).toContain('just one paragraph');
  });

  it('handles nested divs in signature block correctly', () => {
    const html = `
      <p>body</p>
      <div id="Signature">
        outer
        <div>inner nested</div>
        more outer
      </div>
    `;
    const out = extractSignature(html);
    expect(out.heuristic).toBe('div-signature');
    expect(out.signature).toContain('outer');
    expect(out.signature).toContain('inner nested');
    expect(out.signature).toContain('more outer');
  });
});

describe('capture-signature command', () => {
  it('writes signature to ~/.outlook-cli/signature.html by default', async () => {
    const { deps, writeFile } = makeDeps(
      'AAMk-1',
      '<p>body</p><div id="Signature"><b>D.P.</b></div>',
    );
    const result = await run(deps);
    expect(result.path).toBe('/tmp/fake-home/.outlook-cli/signature.html');
    expect(result.heuristic).toBe('div-signature');
    expect(result.signature).toContain('D.P.');
    expect(writeFile).toHaveBeenCalledWith(
      '/tmp/fake-home/.outlook-cli/signature.html',
      expect.stringContaining('D.P.'),
    );
  });

  it('honors --out override', async () => {
    const { deps, writeFile } = makeDeps('AAMk-1', '<div id="Signature">x</div>');
    const result = await run(deps, { out: '/tmp/custom-sig.html' });
    expect(result.path).toBe('/tmp/custom-sig.html');
    expect(writeFile).toHaveBeenCalledWith(
      '/tmp/custom-sig.html',
      expect.any(String),
    );
  });

  it('honors --from-message override (skips listMessagesInFolder)', async () => {
    const { deps, client } = makeDeps(
      'AAMk-latest',
      '<div id="Signature">x</div>',
    );
    await run(deps, { fromMessage: 'AAMk-explicit' });
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
    expect(client.getMessage).toHaveBeenCalledWith('AAMk-explicit', expect.any(Object));
  });

  it('throws UsageError when SentItems is empty', async () => {
    const { deps } = makeDeps(null, '');
    await expect(run(deps)).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError when source body is empty', async () => {
    const { deps, client } = makeDeps('AAMk-1', '');
    (client.getMessage as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      Id: 'AAMk-1',
      Subject: 'empty',
      Body: { ContentType: 'HTML', Content: '' },
    });
    await expect(run(deps)).rejects.toBeInstanceOf(UsageError);
  });

  it('returns extracted signature in result for verification', async () => {
    const { deps } = makeDeps(
      'AAMk-1',
      '<p>body</p><div id="Signature"><b>Καλημέρα — Plessas</b></div>',
    );
    const result = await run(deps);
    expect(result.signature).toContain('Καλημέρα');
    expect(result.signature).toContain('Plessas');
    expect(result.sourceMessageId).toBe('AAMk-1');
  });
});
