// test_scripts/commands-get-thread.spec.ts
//
// Command-level tests for `get-thread`.

import { describe, it, expect, vi } from 'vitest';

import { run } from '../src/commands/get-thread';
import { UsageError } from '../src/commands/list-mail';
import type { OutlookClient } from '../src/http/outlook-client';
import type { MessageSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';
import type { CliConfig } from '../src/config/config';

const CONFIG = {
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
  account: { upn: 'a@b', puid: 'p', tenantId: 't' },
  bearer: {
    token: 'x.y.z',
    expiresAt: '2099-04-21T12:00:00.000Z',
    audience: 'https://outlook.office.com',
    scopes: [],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

function makeMessage(id: string): MessageSummary {
  return {
    Id: id,
    Subject: `s-${id}`,
    ReceivedDateTime: '2026-04-01T00:00:00Z',
    HasAttachments: false,
    IsRead: false,
    WebLink: '',
  };
}

function makeDeps(clientOverrides: Partial<OutlookClient> = {}) {
  const client = {
    get: vi.fn(),
    listMessagesByConversation: vi.fn(),
    ...clientOverrides,
  } as unknown as OutlookClient;
  return {
    deps: {
      config: CONFIG,
      sessionPath: '/tmp/session.json',
      loadSession: vi.fn(async () => SESSION),
      saveSession: vi.fn(async () => {}),
      doAuthCapture: vi.fn(async () => SESSION),
      createClient: vi.fn(() => client),
    },
    client,
  };
}

describe('get-thread', () => {
  it('(message-id mode) fetches the message for ConversationId, then lists by conversation', async () => {
    const { deps, client } = makeDeps();
    (client.get as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      Id: 'AAMk=msg1',
      ConversationId: 'CONV-XYZ',
    });
    (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mockResolvedValueOnce([
      makeMessage('m1'),
      makeMessage('m2'),
    ]);

    const result = await run(deps, 'AAMk=msg1');

    expect(result.conversationId).toBe('CONV-XYZ');
    expect(result.count).toBe(2);
    expect(result.messages.length).toBe(2);
    // First call must be the tight $select lookup.
    const [path, query] = (client.get as ReturnType<typeof vi.fn>).mock.calls[0] as [string, Record<string, string>];
    expect(path).toContain('/api/v2.0/me/messages/');
    expect(query.$select).toBe('Id,ConversationId');
    // Second call goes to listMessagesByConversation.
    expect(client.listMessagesByConversation).toHaveBeenCalledTimes(1);
    const [convId, opts] = (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { orderBy?: string; select?: string[] }];
    expect(convId).toBe('CONV-XYZ');
    expect(opts.orderBy).toBe('ReceivedDateTime asc');
    // Default body = text, so Body + BodyPreview are in the $select list.
    expect(opts.select).toContain('Body');
    expect(opts.select).toContain('BodyPreview');
    expect(opts.select).toContain('ConversationId');
  });

  it('(conv: mode) skips the first GET and queries the conversation directly', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mockResolvedValueOnce([
      makeMessage('m1'),
    ]);

    const result = await run(deps, 'conv:CONV-DIRECT');

    expect(client.get).not.toHaveBeenCalled();
    expect(result.conversationId).toBe('CONV-DIRECT');
    expect(result.count).toBe(1);
    expect(client.listMessagesByConversation).toHaveBeenCalledWith(
      'CONV-DIRECT',
      expect.any(Object),
    );
  });

  it('honors --order desc', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mockResolvedValueOnce([]);
    await run(deps, 'conv:CID', { order: 'desc' });
    const [, opts] = (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { orderBy?: string }];
    expect(opts.orderBy).toBe('ReceivedDateTime desc');
  });

  it('--body none removes Body/BodyPreview from $select', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mockResolvedValueOnce([]);
    await run(deps, 'conv:CID', { body: 'none' });
    const [, opts] = (client.listMessagesByConversation as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { select?: string[] }];
    expect(opts.select).not.toContain('Body');
    expect(opts.select).not.toContain('BodyPreview');
  });

  it('throws UsageError on empty positional', async () => {
    const { deps } = makeDeps();
    await expect(run(deps, '')).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError on empty conv: suffix', async () => {
    const { deps } = makeDeps();
    await expect(run(deps, 'conv:')).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError on invalid --body', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, 'conv:CID', { body: 'wrong' as unknown as 'text' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError on invalid --order', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, 'conv:CID', { order: 'sideways' as unknown as 'asc' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('throws UsageError when message has no ConversationId', async () => {
    const { deps, client } = makeDeps();
    (client.get as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      Id: 'AAMk=msg1',
      // ConversationId deliberately absent
    });
    await expect(run(deps, 'AAMk=msg1')).rejects.toBeInstanceOf(UsageError);
  });
});
