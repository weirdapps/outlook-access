// test_scripts/commands-list-mail-count.spec.ts
//
// Tests the `--just-count` flag on list-mail. The flag should route all three
// folder paths (fast-path alias, --folder-id, resolver) through the new
// `countMessagesInFolder` client method, ignoring --top and --select.

import { describe, it, expect, vi } from 'vitest';

import { run } from '../src/commands/list-mail';
import type { OutlookClient } from '../src/http/outlook-client';
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

function makeDeps(clientOverrides: Partial<OutlookClient> = {}) {
  const client = {
    get: vi.fn(),
    listMessagesInFolder: vi.fn(),
    countMessagesInFolder: vi.fn(),
    getFolder: vi.fn(),
    listFolders: vi.fn(),
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

describe('list-mail --just-count', () => {
  it('(fast-path alias) routes through countMessagesInFolder and returns ListMailCountResult', async () => {
    const { deps, client } = makeDeps();
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 4273,
      exact: true,
    });

    const result = await run(deps, { justCount: true });

    expect(Array.isArray(result)).toBe(false);
    expect(result).toEqual({ count: 4273, exact: true });
    expect(client.countMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId, opts] = (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(folderId).toBe('Inbox');
    expect(opts.filter).toBeUndefined();
    // The list paths must NOT be invoked in count mode.
    expect(client.get).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
  });

  it('(--folder-id) routes through countMessagesInFolder with the raw id verbatim', async () => {
    const { deps, client } = makeDeps();
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 12,
      exact: true,
    });

    await run(deps, { justCount: true, folderId: 'AAMk-raw' });

    const [folderId] = (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, unknown];
    expect(folderId).toBe('AAMk-raw');
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
  });

  it('(resolver path) resolves the folder then counts against the resolved id', async () => {
    const { deps, client } = makeDeps();
    // Resolver walk: MsgFolderRoot → Inbox → Projects
    (client.getFolder as ReturnType<typeof vi.fn>).mockImplementation(
      async (idOrAlias: string) => ({
        Id: `id-${idOrAlias}`,
        DisplayName: idOrAlias,
      }),
    );
    (client.listFolders as ReturnType<typeof vi.fn>).mockResolvedValueOnce([
      { Id: 'id-Projects', DisplayName: 'Projects' },
    ]);
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 7,
      exact: true,
    });

    const result = await run(deps, {
      justCount: true,
      folder: 'Inbox/Projects',
    });

    expect(result).toEqual({ count: 7, exact: true });
    const [folderId] = (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, unknown];
    expect(folderId).toBe('id-Projects');
  });

  it('threads --from/--to into countMessagesInFolder.filter', async () => {
    const { deps, client } = makeDeps();
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 42,
      exact: true,
    });

    await run(deps, {
      justCount: true,
      from: '2026-04-01T00:00:00Z',
      to: '2026-05-01T00:00:00Z',
    });

    const [, opts] = (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mock.calls[0] as [string, { filter?: string }];
    expect(opts.filter).toContain('ReceivedDateTime ge 2026-04-01');
    expect(opts.filter).toContain('ReceivedDateTime lt 2026-05-01');
  });

  it('ignores --top in count mode (no range validation)', async () => {
    const { deps, client } = makeDeps();
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 0,
      exact: true,
    });
    // top: 5000 would normally throw (>1000) — in count mode it's a no-op.
    await expect(
      run(deps, { justCount: true, top: 5000 }),
    ).resolves.toEqual({ count: 0, exact: true });
  });

  it('propagates exact:false when the server did not honor $count=true', async () => {
    const { deps, client } = makeDeps();
    (client.countMessagesInFolder as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      count: 1,
      exact: false,
    });
    const result = await run(deps, { justCount: true });
    expect(result).toEqual({ count: 1, exact: false });
  });
});
