// test_scripts/commands-reply.spec.ts

import { describe, it, expect, vi } from 'vitest';

import { run, composeReplyBody, UsageError } from '../src/commands/reply';
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
    scopes: ['Mail.ReadWrite', 'Mail.Send'],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

const SAMPLE_QUOTED = '<html><body><div>Auto-quoted original from sender.</div></body></html>';

function makeDeps(
  fileMap: Record<string, string> = {},
  draftToRecipients: { EmailAddress: { Address: string } }[] = [
    { EmailAddress: { Address: 'sender@x.com' } },
  ],
) {
  const client = {
    createReply: vi.fn(async (id: string) => ({
      Id: `${id}-reply-draft`,
      WebLink: 'https://outlook/reply',
      Subject: 'RE: original subject',
      ConversationId: 'conv-1',
      Body: { ContentType: 'HTML' as const, Content: SAMPLE_QUOTED },
      ToRecipients: draftToRecipients,
    })),
    createReplyAll: vi.fn(async (id: string) => ({
      Id: `${id}-replyall-draft`,
      WebLink: 'https://outlook/replyall',
      Subject: 'RE: original subject',
      ConversationId: 'conv-1',
      Body: { ContentType: 'HTML' as const, Content: SAMPLE_QUOTED },
      ToRecipients: [
        { EmailAddress: { Address: 'sender@x.com' } },
        { EmailAddress: { Address: 'other@y.com' } },
      ],
    })),
    createForward: vi.fn(async (id: string) => ({
      Id: `${id}-fwd-draft`,
      WebLink: 'https://outlook/fwd',
      Subject: 'FW: original subject',
      ConversationId: 'conv-1',
      Body: { ContentType: 'HTML' as const, Content: SAMPLE_QUOTED },
      ToRecipients: [],
    })),
    updateMessage: vi.fn(async () => ({ Id: 'patched', Subject: 'patched' })),
    sendDraft: vi.fn(async () => undefined),
  } as unknown as OutlookClient;

  const activateOutlook = vi.fn(async () => undefined);
  const readFile = vi.fn(async (p: string) => {
    if (p in fileMap) return Buffer.from(fileMap[p] as string, 'utf-8');
    throw Object.assign(new Error(`ENOENT: ${p}`), { code: 'ENOENT' });
  });

  return {
    deps: {
      config: MINIMAL_CONFIG,
      sessionPath: '/tmp/session.json',
      loadSession: vi.fn(async () => SESSION),
      saveSession: vi.fn(async () => {}),
      doAuthCapture: vi.fn(async () => SESSION),
      createClient: vi.fn(() => client),
      activateOutlook,
      readFile,
      homeDir: () => '/tmp/fake-home',
    },
    client,
    activateOutlook,
    readFile,
  };
}

describe('composeReplyBody', () => {
  it('inserts user content + signature after <body> tag, preserving auto-quote', () => {
    const out = composeReplyBody(
      '<html><body><div>quoted</div></body></html>',
      '<p>my reply</p>',
      '<b>D.P.</b>',
    );
    expect(out).toContain('<body>');
    expect(out.indexOf('my reply')).toBeLessThan(out.indexOf('quoted'));
    expect(out.indexOf('D.P.')).toBeLessThan(out.indexOf('quoted'));
    expect(out).toContain('quoted'); // auto-quote preserved
  });

  it('omits signature block when signature is empty', () => {
    const out = composeReplyBody('<html><body>q</body></html>', '<p>r</p>', '');
    expect(out).toContain('<p>r</p>');
    expect(out).not.toContain('<br><br>');
  });

  it('prepends to whole HTML when no <body> tag present', () => {
    const out = composeReplyBody('<div>quoted</div>', '<p>my</p>', '');
    expect(out.startsWith('<p>my</p>')).toBe(true);
    expect(out).toContain('quoted');
  });
});

describe('reply / reply-all / forward — input validation', () => {
  it('rejects missing source message id', async () => {
    const { deps } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await expect(run(deps, 'reply', '', { html: '/tmp/r.html' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('rejects when neither --html nor --text provided', async () => {
    const { deps } = makeDeps();
    await expect(run(deps, 'reply', 'AAMk-1', {})).rejects.toBeInstanceOf(UsageError);
  });

  it('reply rejects --to (only meaningful for forward)', async () => {
    const { deps } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await expect(
      run(deps, 'reply', 'AAMk-1', { html: '/tmp/r.html', to: 'extra@x.com' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('forward requires --to', async () => {
    const { deps } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await expect(run(deps, 'forward', 'AAMk-1', { html: '/tmp/r.html' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });
});

describe('reply — happy paths', () => {
  it('default → createReply, patches body, activates Outlook, returns draft', async () => {
    const { deps, client, activateOutlook } = makeDeps({
      '/tmp/r.html': '<p>my reply</p>',
      '/tmp/fake-home/.outlook-cli/signature.html': '<b>D.P.</b>',
    });
    const result = await run(deps, 'reply', 'AAMk-source', { html: '/tmp/r.html' });
    expect(client.createReply).toHaveBeenCalledWith('AAMk-source');
    expect(client.updateMessage).toHaveBeenCalledTimes(1);
    expect(client.sendDraft).not.toHaveBeenCalled();
    expect(activateOutlook).toHaveBeenCalledTimes(1);
    expect(result.kind).toBe('reply');
    expect(result.mode).toBe('draft');
    expect(result.id).toBe('AAMk-source-reply-draft');
    expect(result.signatureApplied).toBe(true);
    expect(result.hasQuotedOriginal).toBe(true);
  });

  it('--no-signature suppresses signature lookup', async () => {
    const { deps, readFile } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    const result = await run(deps, 'reply', 'AAMk-1', {
      html: '/tmp/r.html',
      noSignature: true,
    });
    // Body file IS read, but signature file is NOT
    expect(readFile).toHaveBeenCalledTimes(1);
    expect(result.signatureApplied).toBe(false);
  });

  it('--no-open suppresses Outlook activation', async () => {
    const { deps, activateOutlook } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await run(deps, 'reply', 'AAMk-1', { html: '/tmp/r.html', open: false, noSignature: true });
    expect(activateOutlook).not.toHaveBeenCalled();
  });

  it('--send-now → calls sendDraft and returns mode: sent', async () => {
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    const result = await run(deps, 'reply', 'AAMk-1', {
      html: '/tmp/r.html',
      sendNow: true,
      noSignature: true,
    });
    expect(client.sendDraft).toHaveBeenCalledWith('AAMk-1-reply-draft');
    expect(result.mode).toBe('sent');
  });

  it('signature absent file is non-fatal', async () => {
    const { deps } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    // No signature.html in fileMap — should not throw
    const result = await run(deps, 'reply', 'AAMk-1', { html: '/tmp/r.html' });
    expect(result.mode).toBe('draft');
    expect(result.signatureApplied).toBe(false);
  });

  it('--text body wraps plain text in <p> with HTML escaping', async () => {
    const { deps, client } = makeDeps({
      '/tmp/r.txt': 'plain & dangerous <text>',
    });
    await run(deps, 'reply', 'AAMk-1', { text: '/tmp/r.txt', noSignature: true });
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    expect(patch.Body.Content).toContain('plain &amp; dangerous &lt;text&gt;');
  });
});

describe('reply-all', () => {
  it('uses createReplyAll endpoint', async () => {
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await run(deps, 'reply-all', 'AAMk-1', { html: '/tmp/r.html', noSignature: true });
    expect(client.createReplyAll).toHaveBeenCalledWith('AAMk-1');
    expect(client.createReply).not.toHaveBeenCalled();
  });
});

describe('forward', () => {
  it('uses createForward endpoint and patches ToRecipients', async () => {
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>see attached</p>' });
    const result = await run(deps, 'forward', 'AAMk-1', {
      html: '/tmp/r.html',
      to: 'colleague@x.com',
      noSignature: true,
    });
    expect(client.createForward).toHaveBeenCalledWith('AAMk-1');
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    expect(patch.ToRecipients).toEqual([{ EmailAddress: { Address: 'colleague@x.com' } }]);
    expect(result.to).toEqual(['colleague@x.com']);
  });

  it('forward accepts CC and BCC overrides (CC-self default ON adds 1 more)', async () => {
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await run(deps, 'forward', 'AAMk-1', {
      html: '/tmp/r.html',
      to: 'a@x.com',
      cc: 'b@y.com, c@y.com',
      bcc: 'audit@nbg.gr',
      noSignature: true,
    });
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    // 2 user CCs + 1 self-CC (default ON, session.account.upn = me@nbg.gr)
    expect(patch.CcRecipients).toHaveLength(3);
    expect(patch.CcRecipients!.map((r: any) => r.EmailAddress.Address)).toContain('me@nbg.gr');
    expect(patch.BccRecipients).toEqual([{ EmailAddress: { Address: 'audit@nbg.gr' } }]);
  });

  it('forward --no-cc-self (ccSelf: false) suppresses self-CC', async () => {
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await run(deps, 'forward', 'AAMk-1', {
      html: '/tmp/r.html',
      to: 'a@x.com',
      cc: 'b@y.com, c@y.com',
      noSignature: true,
      ccSelf: false,
    });
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    expect(patch.CcRecipients).toHaveLength(2);
    expect(patch.CcRecipients!.map((r: any) => r.EmailAddress.Address)).not.toContain('me@nbg.gr');
  });

  it('reply auto-CCs self when not already in To/Cc (default ON)', async () => {
    // Default mock has draft.ToRecipients = [{ Address: 'sender@x.com' }] only
    const { deps, client } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await run(deps, 'reply', 'AAMk-1', { html: '/tmp/r.html', noSignature: true });
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    expect(patch.CcRecipients).toBeDefined();
    expect(patch.CcRecipients!.map((r: any) => r.EmailAddress.Address)).toContain('me@nbg.gr');
  });

  it('reply ALWAYS adds self to CC even when already in original To (audit/archive)', async () => {
    const { deps, client } = makeDeps(
      { '/tmp/r.html': '<p>r</p>' },
      [{ EmailAddress: { Address: 'me@nbg.gr' } }], // user already in To
    );
    await run(deps, 'reply', 'AAMk-1', { html: '/tmp/r.html', noSignature: true });
    const [, patch] = (client.updateMessage as ReturnType<typeof vi.fn>).mock.calls[0]!;
    // Per CLAUDE.md compliance + user's "ALWAYS cc myself" rule, self goes
    // in CC even when also in TO.
    expect(patch.CcRecipients).toBeDefined();
    expect(patch.CcRecipients!.map((r: any) => r.EmailAddress.Address)).toContain('me@nbg.gr');
  });

  it('forward rejects malformed --to address', async () => {
    const { deps } = makeDeps({ '/tmp/r.html': '<p>r</p>' });
    await expect(
      run(deps, 'forward', 'AAMk-1', { html: '/tmp/r.html', to: 'not-email' }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});
