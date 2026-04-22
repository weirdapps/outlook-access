// test_scripts/commands-send-mail.spec.ts
//
// Unit tests for src/commands/send-mail.ts — recipient parsing, body load,
// attachments, CC-self, dispatch (dry-run/draft/send-now/no-open).

import { describe, it, expect, vi } from 'vitest';

import { run, UsageError } from '../src/commands/send-mail';
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
    scopes: ['Mail.Send'],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

function makeDeps(clientOverrides: Partial<OutlookClient> = {}, fileMap: Record<string, string> = {}) {
  const client = {
    sendMail: vi.fn(async () => undefined),
    createDraft: vi.fn(async () => ({
      Id: 'AAMk-draft-001',
      WebLink: 'https://outlook.office.com/mail/drafts/id/AAMk-draft-001',
      ConversationId: 'conv-001',
    })),
    ...clientOverrides,
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
    },
    client,
    activateOutlook,
    readFile,
  };
}

describe('send-mail — input validation', () => {
  it('rejects when --to is missing', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, { subject: 's', html: '/tmp/b.html' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects when --subject is missing', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, { to: 'a@x.com', html: '/tmp/b.html' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects when neither --html nor --text is provided', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, { to: 'a@x.com', subject: 's' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects malformed recipient address (no @)', async () => {
    const { deps } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await expect(
      run(deps, { to: 'not-an-address', subject: 's', html: '/tmp/b.html' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects when body file does not exist', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, { to: 'a@x.com', subject: 's', html: '/tmp/missing.html' }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});

describe('send-mail — recipient parsing', () => {
  it('accepts comma-separated --to string', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, { to: 'a@x.com, b@y.com', subject: 's', html: '/tmp/b.html' });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.ToRecipients).toEqual([
      { EmailAddress: { Address: 'a@x.com' } },
      { EmailAddress: { Address: 'b@y.com' } },
    ]);
  });

  it('accepts repeated --to as array', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: ['a@x.com', 'b@y.com'],
      subject: 's',
      html: '/tmp/b.html',
    });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.ToRecipients).toHaveLength(2);
  });

  it('mixes comma-string and repeat correctly', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: ['a@x.com, b@y.com', 'c@z.com'],
      subject: 's',
      html: '/tmp/b.html',
    });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.ToRecipients).toHaveLength(3);
  });
});

describe('send-mail — CC-self default', () => {
  it('CC-self ON by default — appends session.account.upn to CC', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, { to: 'a@x.com', subject: 's', html: '/tmp/b.html' });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.CcRecipients).toEqual([{ EmailAddress: { Address: 'me@nbg.gr' } }]);
  });

  it('--no-cc-self (ccSelf: false) suppresses self-CC', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      ccSelf: false,
    });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.CcRecipients).toBeUndefined();
  });

  it('does not double-add self when already in CC (case-insensitive)', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: 'a@x.com',
      cc: 'ME@nbg.gr',
      subject: 's',
      html: '/tmp/b.html',
    });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.CcRecipients).toHaveLength(1);
    expect(payload.CcRecipients![0].EmailAddress.Address).toBe('ME@nbg.gr');
  });
});

describe('send-mail — body & attachments', () => {
  it('--html → ContentType: HTML', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<h1>Καλημέρα</h1>' });
    await run(deps, { to: 'a@x.com', subject: 's', html: '/tmp/b.html' });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.Body.ContentType).toBe('HTML');
    expect(payload.Body.Content).toBe('<h1>Καλημέρα</h1>');
  });

  it('--text → ContentType: Text', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.txt': 'plain text body' });
    await run(deps, { to: 'a@x.com', subject: 's', text: '/tmp/b.txt' });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.Body.ContentType).toBe('Text');
    expect(payload.Body.Content).toBe('plain text body');
  });

  it('attaches files with correct base64 encoding and MIME type', async () => {
    const { deps, client } = makeDeps(
      {},
      {
        '/tmp/b.html': '<p>hi</p>',
        '/tmp/report.pdf': 'PDF-CONTENT',
      },
    );
    await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      attach: ['/tmp/report.pdf'],
    });
    const payload = (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0];
    expect(payload.Attachments).toHaveLength(1);
    const att = payload.Attachments![0];
    expect(att.Name).toBe('report.pdf');
    expect(att.ContentType).toBe('application/pdf');
    expect(Buffer.from(att.ContentBytes, 'base64').toString('utf-8')).toBe('PDF-CONTENT');
    expect(att.IsInline).toBe(false);
  });
});

describe('send-mail — dispatch', () => {
  it('default → calls createDraft and activates Outlook', async () => {
    const { deps, client, activateOutlook } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    const result = await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
    });
    expect(client.createDraft).toHaveBeenCalledTimes(1);
    expect(client.sendMail).not.toHaveBeenCalled();
    expect(activateOutlook).toHaveBeenCalledTimes(1);
    expect(result.mode).toBe('draft');
    expect(result.id).toBe('AAMk-draft-001');
    expect(result.webLink).toContain('outlook.office.com');
  });

  it('--no-open (open: false) → creates draft but skips activation', async () => {
    const { deps, client, activateOutlook } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      open: false,
    });
    expect(client.createDraft).toHaveBeenCalledTimes(1);
    expect(activateOutlook).not.toHaveBeenCalled();
  });

  it('--send-now → calls sendMail (no draft, no activation)', async () => {
    const { deps, client, activateOutlook } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    const result = await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      sendNow: true,
    });
    expect(client.sendMail).toHaveBeenCalledTimes(1);
    expect(client.createDraft).not.toHaveBeenCalled();
    expect(activateOutlook).not.toHaveBeenCalled();
    expect(result.mode).toBe('sent');
  });

  it('--send-now + --no-save-sent → SaveToSentItems false propagated', async () => {
    const { deps, client } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      sendNow: true,
      saveSent: false,
    });
    const [, sendOpts] = (client.sendMail as ReturnType<typeof vi.fn>).mock.calls[0]!;
    expect(sendOpts).toEqual({ saveToSentItems: false });
  });

  it('--dry-run → returns payload, no client calls', async () => {
    const { deps, client, activateOutlook } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    const result = await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
      dryRun: true,
    });
    expect(client.createDraft).not.toHaveBeenCalled();
    expect(client.sendMail).not.toHaveBeenCalled();
    expect(activateOutlook).not.toHaveBeenCalled();
    expect(result.mode).toBe('dry-run');
    expect(result.payload).toBeDefined();
    expect(result.payload!.Subject).toBe('s');
  });

  it('Outlook activation failure is non-fatal — draft still returned', async () => {
    const { deps } = makeDeps({}, { '/tmp/b.html': '<p>hi</p>' });
    deps.activateOutlook = vi.fn(async () => {
      throw new Error('open failed');
    });
    const result = await run(deps, {
      to: 'a@x.com',
      subject: 's',
      html: '/tmp/b.html',
    });
    expect(result.mode).toBe('draft');
    expect(result.id).toBe('AAMk-draft-001');
  });
});
