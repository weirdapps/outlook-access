// test_scripts/commands-list-mail-folder.spec.ts
//
// Unit tests for the Phase-7 folder-flag extension of
// src/commands/list-mail.ts. Exercises the additive `--folder-id` /
// `--folder-parent` flags and their mutual-exclusivity rules (§10.7 of
// project-design.md) against the ACTUAL implementation.
//
// No real HTTP: the test supplies a stub `OutlookClient` via `deps.createClient`
// and asserts on the method(s) the command invokes.

import { describe, expect, it, vi } from 'vitest';

import * as listMail from '../src/commands/list-mail';
import { UsageError } from '../src/commands/list-mail';
import type { CliConfig } from '../src/config/config';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary, MessageSummary, ODataListResponse } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Fakes / fixtures
// ---------------------------------------------------------------------------

const FUTURE_ISO = '2099-04-21T12:00:00.000Z';
const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: {
      upn: 'alice@contoso.com',
      puid: '1234567890',
      tenantId: 'tenant-id-abc',
    },
    bearer: {
      token: JWT_SHAPED_TOKEN,
      expiresAt: FUTURE_ISO,
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read'],
    },
    cookies: [
      {
        name: 'SessionCookie',
        value: 'outlook-cookie-value',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
    ],
    anchorMailbox: 'PUID:1234567890@tenant-id-abc',
  };
}

function buildFakeConfig(overrides: Partial<CliConfig> = {}): CliConfig {
  const base: CliConfig = {
    httpTimeoutMs: 30_000,
    loginTimeoutMs: 300_000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/does-not-exist/session.json',
    profileDir: '/tmp/does-not-exist/profile',
    tz: 'UTC',
    outputMode: 'json',
    listMailTop: 10,
    listMailFolder: 'Inbox',
    bodyMode: 'text',
    calFrom: 'now',
    calTo: 'now + 7d',
    quiet: true,
    noAutoReauth: false,
    // sessionFileOverride / logFilePath intentionally left unset.
    ...overrides,
  };
  return Object.freeze(base);
}

function makeMessage(id: string): MessageSummary {
  return {
    Id: id,
    Subject: `subject-${id}`,
    ReceivedDateTime: '2026-04-20T10:00:00Z',
    HasAttachments: false,
    IsRead: true,
    WebLink: `https://outlook.office.com/mail/${id}`,
    From: { EmailAddress: { Name: 'Alice', Address: 'alice@contoso.com' } },
  };
}

function makeFolder(id: string, displayName: string): FolderSummary {
  return {
    Id: id,
    DisplayName: displayName,
    ParentFolderId: 'parent-of-' + id,
    ChildFolderCount: 0,
    UnreadItemCount: 0,
    TotalItemCount: 0,
    IsHidden: false,
    CreatedDateTime: '2020-01-01T00:00:00Z',
  };
}

/**
 * Minimal `OutlookClient` stub. Every method is a `vi.fn()` whose default
 * implementation fails the test — tests wire in only the methods they expect
 * the command under test to call.
 */
interface StubClient extends OutlookClient {
  get: ReturnType<typeof vi.fn>;
  listFolders: ReturnType<typeof vi.fn>;
  getFolder: ReturnType<typeof vi.fn>;
  createFolder: ReturnType<typeof vi.fn>;
  moveMessage: ReturnType<typeof vi.fn>;
  listMessagesInFolder: ReturnType<typeof vi.fn>;
  listMessagesInFolderAll: ReturnType<typeof vi.fn>;
}

function makeStubClient(): StubClient {
  const stub = {
    get: vi.fn(async () => {
      throw new Error('stub: client.get not configured for this test');
    }),
    listFolders: vi.fn(async () => {
      throw new Error('stub: client.listFolders not configured for this test');
    }),
    getFolder: vi.fn(async () => {
      throw new Error('stub: client.getFolder not configured for this test');
    }),
    createFolder: vi.fn(async () => {
      throw new Error('stub: client.createFolder not configured for this test');
    }),
    moveMessage: vi.fn(async () => {
      throw new Error('stub: client.moveMessage not configured for this test');
    }),
    listMessagesInFolder: vi.fn(async () => {
      throw new Error('stub: client.listMessagesInFolder not configured for this test');
    }),
    listMessagesInFolderAll: vi.fn(async () => {
      throw new Error('stub: client.listMessagesInFolderAll not configured for this test');
    }),
  };
  return stub as unknown as StubClient;
}

function makeDeps(
  overrides: {
    config?: CliConfig;
    client?: StubClient;
    loadSession?: (p: string) => Promise<SessionFile | null>;
    saveSession?: (p: string, s: SessionFile) => Promise<void>;
    doAuthCapture?: () => Promise<SessionFile>;
  } = {},
): { deps: listMail.ListMailDeps; client: StubClient } {
  const client = overrides.client ?? makeStubClient();
  const config = overrides.config ?? buildFakeConfig();
  const session = buildFakeSession();
  const deps: listMail.ListMailDeps = {
    config,
    sessionPath: config.sessionFilePath,
    loadSession: overrides.loadSession ?? (async () => session),
    saveSession:
      overrides.saveSession ??
      (async () => {
        /* no-op */
      }),
    doAuthCapture:
      overrides.doAuthCapture ??
      (async () => {
        throw new Error('doAuthCapture should not be called in these tests');
      }),
    createClient: () => client,
  };
  return { deps, client };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('list-mail folder flag extension (Phase 7)', () => {
  // -----------------------------------------------------------------
  // Default / fast-path branch (no new flag → original behavior preserved)
  // -----------------------------------------------------------------

  it('(1) default (no folder flag) → uses config.listMailFolder via client.listMessagesInFolder', async () => {
    const { deps, client } = makeDeps();
    const messages = [makeMessage('m1'), makeMessage('m2')];
    client.listMessagesInFolder.mockResolvedValueOnce(messages);

    const result = await listMail.run(deps, {});

    expect(result).toEqual(messages);
    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    // Resolver (path) should NOT be taken on the fast path.
    expect(client.listFolders).not.toHaveBeenCalled();
    expect(client.getFolder).not.toHaveBeenCalled();
    expect(client.get).not.toHaveBeenCalled();

    const [folderId, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('Inbox');
    expect(opts).toMatchObject({
      top: 10,
      orderBy: 'ReceivedDateTime desc',
      select: ['Id', 'Subject', 'From', 'ReceivedDateTime', 'HasAttachments', 'IsRead', 'WebLink'],
    });
  });

  it('(2) --folder Archive (well-known alias) → routes via client.listMessagesInFolder with folderId="Archive"', async () => {
    const { deps, client } = makeDeps();
    client.listMessagesInFolder.mockResolvedValueOnce([makeMessage('arch-1')]);

    const result = await listMail.run(deps, { folder: 'Archive' });

    expect(result.map((m) => m.Id)).toEqual(['arch-1']);
    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    expect(client.listFolders).not.toHaveBeenCalled();
    expect(client.getFolder).not.toHaveBeenCalled();
    expect(client.get).not.toHaveBeenCalled();

    const [folderId] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('Archive');
  });

  it('(3) --top and --select are respected on the fast path', async () => {
    const { deps, client } = makeDeps();
    client.listMessagesInFolder.mockResolvedValueOnce([]);

    await listMail.run(deps, {
      top: 25,
      select: 'Id,Subject',
      folder: 'SentItems',
    });

    const [folderId, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('SentItems');
    expect(opts).toMatchObject({
      top: 25,
      orderBy: 'ReceivedDateTime desc',
      select: ['Id', 'Subject'],
    });
  });

  // -----------------------------------------------------------------
  // Path C — resolver branch
  // -----------------------------------------------------------------

  it('(4) --folder "Inbox/Projects" (path in --folder) → resolves via parseFolderSpec + resolveFolder + listMessagesInFolder', async () => {
    const { deps, client } = makeDeps();

    // Well-known-wins-at-root: segment 0 = "Inbox" is resolved via
    // getFolder('Inbox').
    const inboxFolder = makeFolder('inbox-id', 'Inbox');
    const projectsFolder = makeFolder('projects-id', 'Projects');

    client.getFolder.mockImplementation(async (alias: string) => {
      if (alias === 'Inbox') return inboxFolder;
      throw new Error(`unexpected alias ${alias}`);
    });
    client.listFolders.mockImplementation(async (parentId: string) => {
      if (parentId === 'inbox-id') return [projectsFolder];
      throw new Error(`unexpected parentId ${parentId}`);
    });
    client.listMessagesInFolder.mockResolvedValueOnce([makeMessage('p-1')]);

    const result = await listMail.run(deps, { folder: 'Inbox/Projects' });

    expect(result.map((m) => m.Id)).toEqual(['p-1']);

    // Fast path must NOT be taken — client.get must not be called.
    expect(client.get).not.toHaveBeenCalled();

    // Resolver hop must have consulted getFolder + listFolders.
    expect(client.getFolder).toHaveBeenCalledWith('Inbox');
    expect(client.listFolders).toHaveBeenCalledWith('inbox-id');

    // listMessagesInFolder invoked with the resolved leaf id + the merged
    // select/top/orderBy options.
    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('projects-id');
    expect(opts).toMatchObject({
      top: 10,
      orderBy: 'ReceivedDateTime desc',
    });
    expect(opts.select).toEqual([
      'Id',
      'Subject',
      'From',
      'ReceivedDateTime',
      'HasAttachments',
      'IsRead',
      'WebLink',
    ]);
  });

  // -----------------------------------------------------------------
  // Path A — --folder-id (raw id, no resolver hop)
  // -----------------------------------------------------------------

  it('(5) --folder-id AAMk... → uses the id verbatim via client.listMessagesInFolder (no resolver)', async () => {
    const { deps, client } = makeDeps();
    client.listMessagesInFolder.mockResolvedValueOnce([makeMessage('by-id')]);

    const rawId = 'AAMkAGI1234567890Opaque==';
    const result = await listMail.run(deps, { folderId: rawId });

    expect(result.map((m) => m.Id)).toEqual(['by-id']);

    // No resolver hop, no fast path.
    expect(client.get).not.toHaveBeenCalled();
    expect(client.getFolder).not.toHaveBeenCalled();
    expect(client.listFolders).not.toHaveBeenCalled();

    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe(rawId);
    expect(opts).toMatchObject({
      top: 10,
      orderBy: 'ReceivedDateTime desc',
    });
    expect(opts.select).toEqual([
      'Id',
      'Subject',
      'From',
      'ReceivedDateTime',
      'HasAttachments',
      'IsRead',
      'WebLink',
    ]);
  });

  it('(6) --folder-id respects --top and --select overrides', async () => {
    const { deps, client } = makeDeps();
    client.listMessagesInFolder.mockResolvedValueOnce([]);

    await listMail.run(deps, {
      folderId: 'AAMk-raw-id',
      top: 5,
      select: 'Id,Subject,ReceivedDateTime',
    });

    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId, opts] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('AAMk-raw-id');
    expect(opts.top).toBe(5);
    expect(opts.select).toEqual(['Id', 'Subject', 'ReceivedDateTime']);
  });

  // -----------------------------------------------------------------
  // Path C — --folder-parent anchor
  // -----------------------------------------------------------------

  it('(7) --folder-parent Inbox + --folder SubfolderName → resolves SubfolderName under the anchor (path semantics)', async () => {
    const { deps, client } = makeDeps();

    // Because --folder-parent is provided, the command takes the resolver
    // path (path-kind FolderSpec with an attached parent).
    // Parent spec `Inbox` → resolver calls getFolder('Inbox') to resolve
    // the anchor (well-known path). Then walks child 'SubfolderName' under
    // the anchor id via listFolders.
    const inboxFolder = makeFolder('inbox-id', 'Inbox');
    const subFolder = makeFolder('sub-id', 'SubfolderName');

    client.getFolder.mockImplementation(async (alias: string) => {
      if (alias === 'Inbox') return inboxFolder;
      throw new Error(`unexpected alias ${alias}`);
    });
    client.listFolders.mockImplementation(async (parentId: string) => {
      if (parentId === 'inbox-id') return [subFolder];
      throw new Error(`unexpected parentId ${parentId}`);
    });
    client.listMessagesInFolder.mockResolvedValueOnce([makeMessage('sub-1')]);

    const result = await listMail.run(deps, {
      folder: 'SubfolderName',
      folderParent: 'Inbox',
    });

    expect(result.map((m) => m.Id)).toEqual(['sub-1']);

    // Fast path NOT taken.
    expect(client.get).not.toHaveBeenCalled();

    // Resolver consulted.
    expect(client.getFolder).toHaveBeenCalledWith('Inbox');
    expect(client.listFolders).toHaveBeenCalledWith('inbox-id');
    expect(client.listMessagesInFolder).toHaveBeenCalledTimes(1);
    const [folderId] = client.listMessagesInFolder.mock.calls[0];
    expect(folderId).toBe('sub-id');
  });

  // -----------------------------------------------------------------
  // Mutex rules — reflect the ACTUAL Phase-7 implementation.
  // -----------------------------------------------------------------

  it('(8) mutex: --folder X + --folder-id Y → UsageError (exit 2)', async () => {
    const { deps, client } = makeDeps();
    await expect(listMail.run(deps, { folder: 'Inbox', folderId: 'AAMk-raw' })).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof UsageError &&
          err.code === 'BAD_USAGE' &&
          err.exitCode === 2 &&
          /mutually exclusive/i.test(err.message)
        );
      },
    );
    // No client calls should have been made — the mutex check is pre-REST.
    expect(client.get).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
    expect(client.getFolder).not.toHaveBeenCalled();
    expect(client.listFolders).not.toHaveBeenCalled();
  });

  it('(9) mutex: --folder-parent + --folder-id → UsageError (exit 2)', async () => {
    const { deps, client } = makeDeps();
    await expect(
      listMail.run(deps, {
        folderId: 'AAMk-raw',
        folderParent: 'Inbox',
      }),
    ).rejects.toSatisfy((err: unknown) => {
      return (
        err instanceof UsageError &&
        err.code === 'BAD_USAGE' &&
        err.exitCode === 2 &&
        /--folder-parent/.test(err.message) &&
        /--folder-id/.test(err.message)
      );
    });
    expect(client.get).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
  });

  it('(10) mutex: --folder-parent without --folder → UsageError (exit 2)', async () => {
    const { deps, client } = makeDeps();
    await expect(listMail.run(deps, { folderParent: 'Inbox' })).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof UsageError &&
          err.code === 'BAD_USAGE' &&
          err.exitCode === 2 &&
          /--folder-parent requires --folder/i.test(err.message)
        );
      },
    );
    expect(client.get).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
  });

  // -----------------------------------------------------------------
  // Regression: --top bounds still enforced ahead of the new branches
  // -----------------------------------------------------------------

  it('(11) --top out of range still UsageError (cap raised to 1000 in v1.2.0)', async () => {
    const { deps, client } = makeDeps();
    await expect(listMail.run(deps, { top: 0, folderId: 'AAMk-raw' })).rejects.toBeInstanceOf(
      UsageError,
    );
    await expect(listMail.run(deps, { top: 1001, folder: 'Inbox' })).rejects.toBeInstanceOf(
      UsageError,
    );
    // None of the listing methods should be invoked after a bad --top.
    expect(client.get).not.toHaveBeenCalled();
    expect(client.listMessagesInFolder).not.toHaveBeenCalled();
  });
});
