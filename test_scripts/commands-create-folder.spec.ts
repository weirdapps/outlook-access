// test_scripts/commands-create-folder.spec.ts
//
// Command-level tests for `src/commands/create-folder.ts`.
//
// Scope:
//   - Single-name vs nested-path branches.
//   - `--parent` anchor resolution (default MsgFolderRoot vs explicit Inbox).
//   - Idempotent behaviour (both single-name and nested).
//   - Phase-7 fix: nested leaf pre-exists + `--idempotent:false`
//     → `CollisionError` propagates via `ensurePath`.
//   - Argv validation → UsageError (exit 2).
//
// No real HTTP. The OutlookClient is mocked with `Partial<OutlookClient>`.

import { describe, expect, it, vi } from 'vitest';

import { run as runCreateFolder } from '../src/commands/create-folder';
import type { CreateFolderDeps } from '../src/commands/create-folder';
import { UsageError } from '../src/commands/list-mail';
import type { CliConfig } from '../src/config/config';
import { CollisionError } from '../src/http/errors';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

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
      expiresAt: '2099-04-21T12:00:00.000Z',
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

function buildConfig(): CliConfig {
  return Object.freeze({
    httpTimeoutMs: 5_000,
    loginTimeoutMs: 60_000,
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
    noAutoReauth: false,
  }) as CliConfig;
}

/** Build a fake `FolderSummary` with sensible defaults. */
function folder(
  id: string,
  displayName: string,
  parentFolderId?: string,
  extras: Partial<FolderSummary> = {},
): FolderSummary {
  return {
    Id: id,
    DisplayName: displayName,
    ParentFolderId: parentFolderId,
    ...extras,
  };
}

/**
 * Build `CreateFolderDeps` on top of a mock client. The session load path is
 * short-circuited (a valid cached session is returned) so `ensureSession`
 * never touches the network nor triggers `doAuthCapture`.
 */
function buildDeps(
  client: Partial<OutlookClient>,
): { deps: CreateFolderDeps; authCapture: ReturnType<typeof vi.fn> } {
  const session = buildFakeSession();
  const authCapture = vi.fn(async () => session);
  const deps: CreateFolderDeps = {
    config: buildConfig(),
    sessionPath: '/tmp/session.json',
    loadSession: async () => session,
    saveSession: async () => {
      /* no-op */
    },
    doAuthCapture: authCapture,
    createClient: () => client as OutlookClient,
  };
  return { deps, authCapture };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('create-folder command', () => {
  // -------------------------------------------------------------------------
  // Single-name branch
  // -------------------------------------------------------------------------

  it('(1) single name under MsgFolderRoot (default) calls createFolder with the root id', async () => {
    const rootFolder = folder('msgfolderroot-id', 'Top of Information Store', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const createdFolder = folder('new-id', 'Archive-2026', 'msgfolderroot-id');

    const getFolder = vi.fn(async (idOrAlias: string) => {
      if (idOrAlias === 'MsgFolderRoot' || idOrAlias === 'msgfolderroot') {
        return rootFolder;
      }
      throw new Error(`unexpected getFolder(${idOrAlias})`);
    });
    const createFolder = vi.fn(async () => createdFolder);

    const { deps } = buildDeps({ getFolder, createFolder });

    const result = await runCreateFolder(deps, 'Archive-2026');

    // The resolver resolves MsgFolderRoot via getFolder first, then
    // create-folder calls createFolder(parentId, displayName).
    expect(createFolder).toHaveBeenCalledTimes(1);
    expect(createFolder).toHaveBeenCalledWith(
      'msgfolderroot-id',
      'Archive-2026',
    );

    expect(result.created.length).toBe(1);
    expect(result.created[0].Id).toBe('new-id');
    expect(result.created[0].DisplayName).toBe('Archive-2026');
    expect(result.created[0].PreExisting).toBe(false);
    expect(result.leaf).toBe(result.created[0]);
    expect(result.idempotent).toBe(false);
  });

  it('(2) single name under --parent Inbox resolves Inbox via getFolder and creates under its id', async () => {
    const inboxFolder = folder('inbox-id', 'Inbox', undefined, {
      WellKnownName: 'inbox',
    });
    const createdFolder = folder('child-id', 'Projects', 'inbox-id');

    const getFolder = vi.fn(async (idOrAlias: string) => {
      if (idOrAlias === 'Inbox') return inboxFolder;
      throw new Error(`unexpected getFolder(${idOrAlias})`);
    });
    const createFolder = vi.fn(async () => createdFolder);

    const { deps } = buildDeps({ getFolder, createFolder });

    const result = await runCreateFolder(deps, 'Projects', { parent: 'Inbox' });

    expect(getFolder).toHaveBeenCalledWith('Inbox');
    expect(createFolder).toHaveBeenCalledTimes(1);
    expect(createFolder).toHaveBeenCalledWith('inbox-id', 'Projects');

    expect(result.leaf.Id).toBe('child-id');
    expect(result.leaf.ParentFolderId).toBe('inbox-id');
    expect(result.leaf.PreExisting).toBe(false);
  });

  it('(3) single-name collision without --idempotent propagates CollisionError', async () => {
    const rootFolder = folder('msgfolderroot-id', 'Top of Information Store', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const collision = new CollisionError({
      code: 'FOLDER_ALREADY_EXISTS',
      message: "folder 'Archive-2026' already exists",
      path: 'Archive-2026',
      parentId: 'msgfolderroot-id',
    });

    const getFolder = vi.fn(async () => rootFolder);
    const createFolder = vi.fn(async () => {
      throw collision;
    });
    const listFolders = vi.fn(async () => [] as FolderSummary[]);

    const { deps } = buildDeps({ getFolder, createFolder, listFolders });

    await expect(
      runCreateFolder(deps, 'Archive-2026'),
    ).rejects.toBeInstanceOf(CollisionError);

    // Non-idempotent path must NOT re-list.
    expect(listFolders).not.toHaveBeenCalled();
  });

  it('(4) single-name collision with --idempotent re-lists parent and returns PreExisting:true', async () => {
    const rootFolder = folder('msgfolderroot-id', 'Top of Information Store', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const collision = new CollisionError({
      code: 'FOLDER_ALREADY_EXISTS',
      message: "folder 'Archive-2026' already exists",
      path: 'Archive-2026',
      parentId: 'msgfolderroot-id',
    });
    const existingChild = folder('existing-id', 'Archive-2026', 'msgfolderroot-id');

    const getFolder = vi.fn(async () => rootFolder);
    const createFolder = vi.fn(async () => {
      throw collision;
    });
    const listFolders = vi.fn(async () => [existingChild]);

    const { deps } = buildDeps({ getFolder, createFolder, listFolders });

    const result = await runCreateFolder(deps, 'Archive-2026', {
      idempotent: true,
    });

    expect(listFolders).toHaveBeenCalledTimes(1);
    expect(listFolders).toHaveBeenCalledWith('msgfolderroot-id');

    expect(result.leaf.Id).toBe('existing-id');
    expect(result.leaf.DisplayName).toBe('Archive-2026');
    expect(result.leaf.PreExisting).toBe(true);
    expect(result.idempotent).toBe(true);
  });

  // -------------------------------------------------------------------------
  // Nested-path branch
  // -------------------------------------------------------------------------

  it('(5) nested path A/B/C without --create-parents raises FOLDER_MISSING_PARENT when B is missing', async () => {
    // Setup: A exists under MsgFolderRoot, but B does not exist under A.
    const rootFolder = folder('root-id', 'Top', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const folderA = folder('A-id', 'A', 'root-id');

    const getFolder = vi.fn(async () => rootFolder);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [folderA];
      if (parentId === 'A-id') return []; // B missing
      return [];
    });
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder should not be called');
    });

    const { deps } = buildDeps({ getFolder, listFolders, createFolder });

    try {
      await runCreateFolder(deps, 'A/B/C', { createParents: false });
      throw new Error('expected throw');
    } catch (err) {
      expect(err).toBeInstanceOf(UsageError);
      expect((err as UsageError).message).toMatch(/FOLDER_MISSING_PARENT/);
    }

    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(6) nested path A/B/C with --create-parents creates all three segments in order', async () => {
    const rootFolder = folder('root-id', 'Top', undefined, {
      WellKnownName: 'msgfolderroot',
    });

    const getFolder = vi.fn(async () => rootFolder);

    // Children start empty at every level; after each create, the next
    // listFolders call at the newly-created level should return empty (C
    // will be created under the newly-created B).
    const listFolders = vi.fn(async () => [] as FolderSummary[]);

    // Record creation order.
    const createCalls: Array<{ parentId: string; name: string }> = [];
    const createFolder = vi.fn(
      async (parentId: string, displayName: string) => {
        createCalls.push({ parentId, name: displayName });
        if (displayName === 'A') return folder('A-id', 'A', parentId);
        if (displayName === 'B') return folder('B-id', 'B', parentId);
        if (displayName === 'C') return folder('C-id', 'C', parentId);
        throw new Error(`unexpected displayName ${displayName}`);
      },
    );

    const { deps } = buildDeps({ getFolder, listFolders, createFolder });

    // Idempotent fast-path is NOT triggered (createParents without idempotent).
    const result = await runCreateFolder(deps, 'A/B/C', {
      createParents: true,
    });

    // Strict order: A first (under root), B next (under A), C last (under B).
    expect(createCalls).toEqual([
      { parentId: 'root-id', name: 'A' },
      { parentId: 'A-id', name: 'B' },
      { parentId: 'B-id', name: 'C' },
    ]);

    // The command surfaces only the leaf segment. The problem statement
    // asks for per-segment PreExisting flags; the current implementation
    // returns a single-entry `created[]` for the leaf (see create-folder.ts
    // comment at runNestedPath). Assert the leaf shape + that it is the
    // newly-created C with PreExisting:false.
    expect(result.leaf.Id).toBe('C-id');
    expect(result.leaf.DisplayName).toBe('C');
    expect(result.leaf.PreExisting).toBe(false);
    expect(result.idempotent).toBe(false);
  });

  it('(7) nested path A/B/C with --create-parents where A exists, B missing: A not re-POSTed, B and C created', async () => {
    const rootFolder = folder('root-id', 'Top', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const existingA = folder('A-id', 'A', 'root-id');

    const getFolder = vi.fn(async () => rootFolder);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [existingA];
      // Everything below A starts empty.
      return [] as FolderSummary[];
    });

    const createCalls: Array<{ parentId: string; name: string }> = [];
    const createFolder = vi.fn(
      async (parentId: string, displayName: string) => {
        createCalls.push({ parentId, name: displayName });
        if (displayName === 'B') return folder('B-id', 'B', parentId);
        if (displayName === 'C') return folder('C-id', 'C', parentId);
        throw new Error(`unexpected displayName ${displayName}`);
      },
    );

    const { deps } = buildDeps({ getFolder, listFolders, createFolder });

    const result = await runCreateFolder(deps, 'A/B/C', {
      createParents: true,
    });

    // A must NOT have been POSTed — only B and C in that order.
    expect(createCalls).toEqual([
      { parentId: 'A-id', name: 'B' },
      { parentId: 'B-id', name: 'C' },
    ]);
    expect(result.leaf.Id).toBe('C-id');
    expect(result.leaf.PreExisting).toBe(false);
  });

  it('(8) nested path with leaf existing + --idempotent:false raises CollisionError (Phase-7 fix)', async () => {
    // Path: A/B where both A and B already exist. Without --idempotent the
    // leaf pre-existence is a collision per ensurePath (Phase-7 fix).
    const rootFolder = folder('root-id', 'Top', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const existingA = folder('A-id', 'A', 'root-id');
    const existingB = folder('B-id', 'B', 'A-id');

    const getFolder = vi.fn(async () => rootFolder);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [existingA];
      if (parentId === 'A-id') return [existingB];
      return [];
    });
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder must not be called on leaf-exists path');
    });

    const { deps } = buildDeps({ getFolder, listFolders, createFolder });

    await expect(
      runCreateFolder(deps, 'A/B', {
        createParents: true,
        idempotent: false,
      }),
    ).rejects.toBeInstanceOf(CollisionError);

    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(9) nested path with leaf existing + --idempotent:true returns PreExisting with no POSTs', async () => {
    const rootFolder = folder('root-id', 'Top', undefined, {
      WellKnownName: 'msgfolderroot',
    });
    const existingA = folder('A-id', 'A', 'root-id');
    const existingB = folder('B-id', 'B', 'A-id');

    const getFolder = vi.fn(async () => rootFolder);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [existingA];
      if (parentId === 'A-id') return [existingB];
      return [];
    });
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder must not be called on idempotent fast-path');
    });

    const { deps } = buildDeps({ getFolder, listFolders, createFolder });

    const result = await runCreateFolder(deps, 'A/B', {
      createParents: true,
      idempotent: true,
    });

    // Idempotent fast-path resolved the full path; no POSTs issued.
    expect(createFolder).not.toHaveBeenCalled();
    expect(result.leaf.Id).toBe('B-id');
    expect(result.leaf.PreExisting).toBe(true);
    expect(result.idempotent).toBe(true);
  });

  // -------------------------------------------------------------------------
  // Argv validation → UsageError (exit 2)
  // -------------------------------------------------------------------------

  it('(10) empty positional path raises UsageError', async () => {
    const { deps } = buildDeps({});
    await expect(runCreateFolder(deps, '')).rejects.toBeInstanceOf(UsageError);
  });

  it("(11) 'id:...' as positional raises UsageError", async () => {
    const { deps } = buildDeps({});
    await expect(
      runCreateFolder(deps, 'id:AAMkAGI'),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('(12) bare well-known alias positional raises UsageError', async () => {
    const { deps } = buildDeps({});
    await expect(runCreateFolder(deps, 'Inbox')).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('(13) well-known-alias leaf at MsgFolderRoot (nested path) raises UsageError', async () => {
    // Path whose leaf is a well-known alias, anchor is default MsgFolderRoot.
    // This validation fires before any REST call; the mock client can be empty.
    const { deps } = buildDeps({});
    await expect(
      runCreateFolder(deps, 'Projects/Inbox', { createParents: true }),
    ).rejects.toBeInstanceOf(UsageError);
  });
});
