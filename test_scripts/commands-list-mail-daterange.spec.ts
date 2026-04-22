// test_scripts/commands-list-mail-daterange.spec.ts
//
// Tests the `--from` / `--to` extension of `list-mail`.

import { describe, it, expect, vi } from 'vitest';

import { run, UsageError } from '../src/commands/list-mail';
import type { OutlookClient } from '../src/http/outlook-client';
import type { MessageSummary, ODataListResponse } from '../src/http/types';
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
    Subject: id,
    ReceivedDateTime: '2026-04-01T00:00:00Z',
    HasAttachments: false,
    IsRead: false,
    WebLink: '',
  };
}

function makeDeps(clientOverrides: Partial<OutlookClient> = {}) {
  const client = {
    get: vi.fn(),
    listMessagesInFolder: vi.fn(),
    ...clientOverrides,
  } as unknown as OutlookClient;
  return {
    deps: {
      config: MINIMAL_CONFIG,
      sessionPath: '/tmp/session.json',
      loadSession: vi.fn(async () => SESSION),
      saveSession: vi.fn(async () => {}),
      doAuthCapture: vi.fn(async () => SESSION),
      createClient: vi.fn(() => client),
    },
    client,
  };
}

// Architecture note for fork: list-mail.ts in this fork routes ALL paths
// through `client.listMessagesInFolder(folderId, opts)`. Upstream's v1.2.0
// uses `client.get(path, query)` directly for the fast-path alias case;
// these tests are adapted to mock listMessagesInFolder for fast-path too.
describe('list-mail --from / --to', () => {
  it('(fast path) builds $filter=ReceivedDateTime ge X and lt Y when both bounds are set', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce([
      makeMessage('m1'),
    ]);

    const from = '2026-04-01T00:00:00Z';
    const to = '2026-05-01T00:00:00Z';
    await run(deps, { top: 5, from, to });

    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId, listOpts] = (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(folderId).toBe('Inbox');
    expect(listOpts.filter).toContain('ReceivedDateTime ge 2026-04-01');
    expect(listOpts.filter).toContain('ReceivedDateTime lt 2026-05-01');
    expect(listOpts.filter).toContain(' and ');
  });

  it('(fast path) builds only lower bound when only --from is set', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce([]);
    await run(deps, { from: 'now - 1d' });
    const [, listOpts] = (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(listOpts.filter).toMatch(/^ReceivedDateTime ge \S+$/);
    expect(listOpts.filter).not.toContain(' and ');
  });

  it('(fast path) omits $filter entirely when neither bound is set', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce([]);
    await run(deps, {});
    const [, listOpts] = (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(listOpts.filter).toBeUndefined();
  });

  it('(--folder-id path) threads filter into listMessagesInFolder.filter', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce([
      makeMessage('m1'),
    ]);
    await run(deps, {
      folderId: 'AAMk=abc',
      from: '2026-04-01T00:00:00Z',
      to: '2026-05-01T00:00:00Z',
    });
    const [folderId, opts] = (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(folderId).toBe('AAMk=abc');
    expect(opts.filter).toContain('ReceivedDateTime ge 2026-04-01');
    expect(opts.filter).toContain('ReceivedDateTime lt 2026-05-01');
  });

  it('accepts "now - Nd" keyword for --from', async () => {
    const { deps, client } = makeDeps();
    (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce([]);
    await run(deps, { from: 'now - 7d' });
    const [, listOpts] = (client.listMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    const match = (listOpts.filter ?? '').match(/ge (\S+)/);
    expect(match).not.toBeNull();
    const iso = match![1]!;
    expect(Number.isFinite(Date.parse(iso))).toBe(true);
  });

  it('rejects malformed --from with UsageError', async () => {
    const { deps } = makeDeps();
    await expect(run(deps, { from: 'not-a-date' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('rejects malformed --to with UsageError', async () => {
    const { deps } = makeDeps();
    await expect(run(deps, { to: '!!garbage!!' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('rejects --since combined with --from (fork-only mutual exclusion)', async () => {
    const { deps } = makeDeps();
    await expect(
      run(deps, { since: '2026-04-01T00:00:00Z', from: '2026-04-01T00:00:00Z' }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});
