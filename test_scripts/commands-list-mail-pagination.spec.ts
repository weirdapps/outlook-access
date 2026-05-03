// test_scripts/commands-list-mail-pagination.spec.ts
//
// Unit tests for the --since/--until/--all/--max extension of list-mail.
// Asserts:
//   - filter is built and passed through to client.listMessagesInFolder
//   - --all routes to client.listMessagesInFolderAll with correct maxResults
//   - --max validation
//   - --since validation
//   - truncated stderr warning

import { describe, expect, it, vi } from 'vitest';

import * as listMail from '../src/commands/list-mail';
import { UsageError } from '../src/commands/list-mail';
import type { CliConfig } from '../src/config/config';
import type { OutlookClient } from '../src/http/outlook-client';
import type { MessageSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

const FUTURE_ISO = '2099-04-21T12:00:00.000Z';
const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: { upn: 'a@b.com', puid: '1', tenantId: 't' },
    bearer: {
      token: JWT_SHAPED_TOKEN,
      expiresAt: FUTURE_ISO,
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read'],
    },
    cookies: [
      {
        name: 'C',
        value: 'v',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
    ],
    anchorMailbox: 'PUID:1@t',
  };
}

function buildFakeConfig(): CliConfig {
  return {
    httpTimeoutMs: 30_000,
    loginTimeoutMs: 300_000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/x',
    profileDir: '/tmp/p',
    tz: 'UTC',
    outputMode: 'json',
    listMailTop: 10,
    listMailFolder: 'Inbox',
    bodyMode: 'text',
    calFrom: 'now',
    calTo: 'now + 7d',
    quiet: true,
    noAutoReauth: false,
  };
}

function makeMessage(id: string): MessageSummary {
  return {
    Id: id,
    Subject: 'x',
    From: { EmailAddress: { Name: 'n', Address: 'a@b' } },
    ReceivedDateTime: '2026-04-22T00:00:00Z',
    HasAttachments: false,
    IsRead: false,
    WebLink: 'https://example.com',
  } as unknown as MessageSummary;
}

interface StubClient extends OutlookClient {
  get: ReturnType<typeof vi.fn>;
  listFolders: ReturnType<typeof vi.fn>;
  getFolder: ReturnType<typeof vi.fn>;
  createFolder: ReturnType<typeof vi.fn>;
  moveMessage: ReturnType<typeof vi.fn>;
  listMessagesInFolder: ReturnType<typeof vi.fn>;
  listMessagesInFolderAll: ReturnType<typeof vi.fn>;
}

function makeClient(): StubClient {
  return {
    get: vi.fn(),
    listFolders: vi.fn(),
    getFolder: vi.fn(),
    createFolder: vi.fn(),
    moveMessage: vi.fn(),
    listMessagesInFolder: vi.fn(),
    listMessagesInFolderAll: vi.fn(),
  } as unknown as StubClient;
}

function makeDeps(client: StubClient): listMail.ListMailDeps {
  const session = buildFakeSession();
  return {
    config: buildFakeConfig(),
    sessionPath: '/tmp/x',
    loadSession: async () => session,
    saveSession: async () => {},
    doAuthCapture: async () => {
      throw new Error('not used');
    },
    createClient: () => client,
  };
}

describe('list-mail --since / --until / --all / --max', () => {
  it('passes filter through to listMessagesInFolder when both bounds set', async () => {
    const client = makeClient();
    client.listMessagesInFolder.mockResolvedValueOnce([makeMessage('m1')]);
    await listMail.run(makeDeps(client), {
      since: '2026-04-22T07:00:00Z',
      until: '2026-04-23T00:00:00Z',
    });
    const [, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(opts.filter).toBe(
      'ReceivedDateTime ge 2026-04-22T07:00:00Z and ReceivedDateTime lt 2026-04-23T00:00:00Z',
    );
  });

  it('passes since-only filter', async () => {
    const client = makeClient();
    client.listMessagesInFolder.mockResolvedValueOnce([]);
    await listMail.run(makeDeps(client), { since: '2026-04-22T00:00:00Z' });
    const [, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(opts.filter).toBe('ReceivedDateTime ge 2026-04-22T00:00:00Z');
  });

  it('omits filter when neither bound set', async () => {
    const client = makeClient();
    client.listMessagesInFolder.mockResolvedValueOnce([]);
    await listMail.run(makeDeps(client), {});
    const [, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(opts.filter).toBeUndefined();
  });

  it('routes to listMessagesInFolderAll when --all is true', async () => {
    const client = makeClient();
    client.listMessagesInFolderAll.mockResolvedValueOnce({
      messages: [makeMessage('m1'), makeMessage('m2')],
      truncated: false,
    });
    const result = await listMail.run(makeDeps(client), { all: true, max: 500 });
    expect(result.length).toBe(2);
    expect(client.listMessagesInFolderAll).toHaveBeenCalledTimes(1);
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
    const [, , maxResults] = client.listMessagesInFolderAll.mock.calls[0];
    expect(maxResults).toBe(500);
  });

  it('uses default max=10000 when --all is set without --max', async () => {
    const client = makeClient();
    client.listMessagesInFolderAll.mockResolvedValueOnce({ messages: [], truncated: false });
    await listMail.run(makeDeps(client), { all: true });
    const [, , maxResults] = client.listMessagesInFolderAll.mock.calls[0];
    expect(maxResults).toBe(10_000);
  });

  it('emits stderr warning when truncated', async () => {
    const stderr = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    const client = makeClient();
    client.listMessagesInFolderAll.mockResolvedValueOnce({
      messages: [makeMessage('m1')],
      truncated: true,
    });
    await listMail.run(makeDeps(client), { all: true, max: 1 });
    const stderrCall = stderr.mock.calls.find((c) => String(c[0]).includes('max_results_reached'));
    expect(stderrCall).toBeDefined();
    stderr.mockRestore();
  });

  it('rejects --max < 1', async () => {
    const client = makeClient();
    await expect(listMail.run(makeDeps(client), { all: true, max: 0 })).rejects.toThrow(UsageError);
    await expect(listMail.run(makeDeps(client), { all: true, max: 0 })).rejects.toThrow(
      /positive integer/,
    );
  });

  it('rejects --max > 100000', async () => {
    const client = makeClient();
    await expect(listMail.run(makeDeps(client), { all: true, max: 100_001 })).rejects.toThrow(
      /cannot exceed 100000/,
    );
  });

  it('rejects --since with malformed timestamp', async () => {
    const client = makeClient();
    await expect(listMail.run(makeDeps(client), { since: 'yesterday' })).rejects.toThrow(
      UsageError,
    );
    await expect(listMail.run(makeDeps(client), { since: 'yesterday' })).rejects.toThrow(
      /ISO-8601/,
    );
  });

  it('rejects since >= until', async () => {
    const client = makeClient();
    await expect(
      listMail.run(makeDeps(client), {
        since: '2026-04-23T00:00:00Z',
        until: '2026-04-22T00:00:00Z',
      }),
    ).rejects.toThrow(/earlier than/);
  });

  it('does not invoke listMessagesInFolderAll when --all is false', async () => {
    const client = makeClient();
    client.listMessagesInFolder.mockResolvedValueOnce([]);
    await listMail.run(makeDeps(client), { all: false });
    expect(client.listMessagesInFolderAll).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
  });
});
