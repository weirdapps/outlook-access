// test_scripts/commands-list-folders.spec.ts
//
// Command-level tests for `src/commands/list-folders.ts` — `run(deps, opts)`
// plus the exported LIST_FOLDERS_COLUMNS descriptor. No real HTTP; the
// OutlookClient is injected as a Partial<OutlookClient> cast. Session loading
// uses a stub loadSession so ensureSession() never touches disk.
//
// Sources of truth:
//   - src/commands/list-folders.ts (ListFoldersDeps, ListFoldersOptions, run,
//     LIST_FOLDERS_COLUMNS)
//   - docs/design/project-design.md §10.7 (flag semantics)
//   - docs/design/project-design.md §10.5 (path grammar / recursive cap)

import { describe, expect, it, vi } from 'vitest';

import type { CliConfig } from '../src/config/config';
import { UpstreamError } from '../src/config/errors';
import { UsageError } from '../src/commands/list-mail';
import {
  LIST_FOLDERS_COLUMNS,
  run as runListFolders,
  type ListFoldersDeps,
  type ListFoldersOptions,
} from '../src/commands/list-folders';
import { MAX_FOLDERS_VISITED } from '../src/folders/types';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

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
      token: 'aaaaaaaa.bbbbbbbb.cccccccc',
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read'],
    },
    cookies: [],
    anchorMailbox: 'PUID:1234567890@tenant-id-abc',
  };
}

function buildFakeConfig(): CliConfig {
  return {
    httpTimeoutMs: 5000,
    loginTimeoutMs: 60000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/never-touched.json',
    profileDir: '/tmp/never-touched-profile',
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

interface BuiltDeps {
  deps: ListFoldersDeps;
  loadSession: ReturnType<typeof vi.fn>;
  saveSession: ReturnType<typeof vi.fn>;
  doAuthCapture: ReturnType<typeof vi.fn>;
  createClient: ReturnType<typeof vi.fn>;
  client: Partial<OutlookClient>;
}

/**
 * Builds a deps object with all side-effects stubbed. The caller supplies a
 * Partial<OutlookClient> (cast internally) that models listFolders + getFolder
 * for the scenario under test. The default session is non-expired so
 * ensureSession() returns it without calling doAuthCapture / saveSession.
 */
function buildDeps(clientOverrides: Partial<OutlookClient> = {}): BuiltDeps {
  const client: Partial<OutlookClient> = { ...clientOverrides };
  const loadSession = vi.fn(async () => buildFakeSession());
  const saveSession = vi.fn(async () => undefined);
  const doAuthCapture = vi.fn(async () => buildFakeSession());
  const createClient = vi.fn(() => client as OutlookClient);

  const deps: ListFoldersDeps = {
    config: buildFakeConfig(),
    sessionPath: '/tmp/never-touched.json',
    loadSession,
    saveSession,
    doAuthCapture,
    createClient,
  };

  return { deps, loadSession, saveSession, doAuthCapture, createClient, client };
}

function folder(
  overrides: Partial<FolderSummary> & { Id: string; DisplayName: string },
): FolderSummary {
  return {
    Id: overrides.Id,
    DisplayName: overrides.DisplayName,
    ...overrides,
  } as FolderSummary;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('list-folders: run()', () => {
  it('(1) default (no --parent) uses MsgFolderRoot directly without a resolver REST hop', async () => {
    const listFolders = vi.fn(async () => [
      folder({ Id: 'f-1', DisplayName: 'Inbox' }),
      folder({ Id: 'f-2', DisplayName: 'Drafts' }),
    ]);
    const getFolder = vi.fn(async () => {
      throw new Error(
        'getFolder must not be called when --parent is unset (MsgFolderRoot short-circuits)',
      );
    });
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
      getFolder: getFolder as OutlookClient['getFolder'],
    });

    const rows = await runListFolders(deps, {});

    expect(getFolder).not.toHaveBeenCalled();
    expect(listFolders).toHaveBeenCalledTimes(1);
    const [parentIdArg, topArg] = listFolders.mock.calls[0];
    expect(parentIdArg).toBe('MsgFolderRoot');
    // Default top is DEFAULT_LIST_FOLDERS_TOP (100).
    expect(topArg).toBe(100);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ Id: 'f-1', DisplayName: 'Inbox', Path: 'Inbox' });
    expect(rows[1]).toMatchObject({ Id: 'f-2', DisplayName: 'Drafts', Path: 'Drafts' });
  });

  it('(2) --parent Inbox resolves the alias via client.getFolder and calls listFolders with the resolved id', async () => {
    // `Inbox` is parsed as { kind: "wellKnown" } and resolved via a single
    // getFolder() call; the returned Id is then passed to listFolders.
    // Only `MsgFolderRoot` short-circuits without a REST hop (ADR-15).
    const listFolders = vi.fn(async () => [folder({ Id: 'child-1', DisplayName: 'Projects' })]);
    const getFolder = vi.fn(async () =>
      folder({ Id: 'inbox-real-id', DisplayName: 'Inbox', WellKnownName: 'inbox' }),
    );
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
      getFolder: getFolder as OutlookClient['getFolder'],
    });

    const rows = await runListFolders(deps, { parent: 'Inbox' });

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('Inbox');
    expect(listFolders).toHaveBeenCalledTimes(1);
    expect(listFolders.mock.calls[0][0]).toBe('inbox-real-id');
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ Id: 'child-1', Path: 'Projects' });
  });

  it('(3) --parent "Inbox/Projects" resolves via the path walk and passes resolved id to listFolders', async () => {
    // Resolver walk for "Inbox/Projects":
    //   segment 0 "Inbox" hits the well-known shortcut → getFolder("Inbox")
    //   segment 1 "Projects" → listFolders(inboxId) then pick the single match
    //   finally, the command calls listFolders(projectsId) to enumerate kids.
    const inboxId = 'inbox-id-xyz';
    const projectsId = 'projects-id-abc';

    const getFolder = vi.fn(async (alias: string) => {
      if (alias === 'Inbox') {
        return folder({ Id: inboxId, DisplayName: 'Inbox', WellKnownName: 'inbox' });
      }
      throw new Error(`unexpected getFolder(${alias})`);
    });

    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === inboxId) {
        return [folder({ Id: projectsId, DisplayName: 'Projects', ChildFolderCount: 2 })];
      }
      if (parentId === projectsId) {
        return [
          folder({ Id: 'p-child-1', DisplayName: 'Alpha' }),
          folder({ Id: 'p-child-2', DisplayName: 'Beta' }),
        ];
      }
      throw new Error(`unexpected listFolders(${parentId})`);
    });

    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
      getFolder: getFolder as OutlookClient['getFolder'],
    });

    const rows = await runListFolders(deps, { parent: 'Inbox/Projects' });

    // getFolder called exactly for the well-known root segment.
    expect(getFolder).toHaveBeenCalledWith('Inbox');
    // listFolders was called twice: once by the resolver, once to enumerate.
    const calls = listFolders.mock.calls.map((c) => c[0]);
    expect(calls).toContain(inboxId);
    expect(calls).toContain(projectsId);
    expect(rows.map((r) => r.Id).sort()).toEqual(['p-child-1', 'p-child-2']);
  });

  it('(4) --top 50 is forwarded to client.listFolders', async () => {
    const listFolders = vi.fn(async () => []);
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    await runListFolders(deps, { top: 50 });

    expect(listFolders).toHaveBeenCalledTimes(1);
    expect(listFolders.mock.calls[0][1]).toBe(50);
  });

  it('(5) --top 300 (>250) rejects with UsageError before any REST call', async () => {
    const listFolders = vi.fn(async () => []);
    const getFolder = vi.fn(async () => folder({ Id: 'x', DisplayName: 'x' }));
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
      getFolder: getFolder as OutlookClient['getFolder'],
    });

    await expect(runListFolders(deps, { top: 300 })).rejects.toBeInstanceOf(UsageError);
    expect(listFolders).not.toHaveBeenCalled();
    expect(getFolder).not.toHaveBeenCalled();
  });

  it('(5b) --top below range (0) also rejects with UsageError', async () => {
    const listFolders = vi.fn(async () => []);
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    await expect(runListFolders(deps, { top: 0 })).rejects.toBeInstanceOf(UsageError);
    expect(listFolders).not.toHaveBeenCalled();
  });

  it('(6) --recursive performs a breadth-first walk and returns Depth + Path on each row', async () => {
    // Tree under MsgFolderRoot:
    //   A  (depth 0, Id=a, ChildFolderCount=1)
    //     A1 (depth 1, Id=a1)
    //   B  (depth 0, Id=b, ChildFolderCount=0)
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'MsgFolderRoot') {
        return [
          folder({ Id: 'a', DisplayName: 'A', ChildFolderCount: 1 }),
          folder({ Id: 'b', DisplayName: 'B', ChildFolderCount: 0 }),
        ];
      }
      if (parentId === 'a') {
        return [folder({ Id: 'a1', DisplayName: 'A1', ChildFolderCount: 0 })];
      }
      return [];
    });

    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const rows = await runListFolders(deps, { recursive: true });

    // Visit order (BFS): root's children first (A, B), then A's children (A1).
    expect(rows).toHaveLength(3);
    expect(rows[0]).toMatchObject({ Id: 'a', Depth: 0, Path: 'A' });
    expect(rows[1]).toMatchObject({ Id: 'b', Depth: 0, Path: 'B' });
    expect(rows[2]).toMatchObject({ Id: 'a1', Depth: 1, Path: 'A/A1' });
  });

  it('(7) --recursive exceeding MAX_FOLDERS_VISITED raises UpstreamError UPSTREAM_PAGINATION_LIMIT', async () => {
    // Return one level of children that is strictly larger than the cap.
    // Each child has ChildFolderCount: 0 so we do NOT descend — the cap is
    // exclusively triggered by accumulating rows.
    const oversized: FolderSummary[] = Array.from({ length: MAX_FOLDERS_VISITED + 1 }, (_v, i) =>
      folder({
        Id: `big-${i}`,
        DisplayName: `F${i}`,
        ChildFolderCount: 0,
      }),
    );
    const listFolders = vi.fn(async () => oversized);

    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    let caught: unknown = null;
    try {
      await runListFolders(deps, { recursive: true });
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(UpstreamError);
    expect((caught as UpstreamError).code).toBe('UPSTREAM_PAGINATION_LIMIT');
    expect((caught as UpstreamError).exitCode).toBe(5);
  });

  it('(8a) --include-hidden off (default) filters out folders whose IsHidden === true', async () => {
    const listFolders = vi.fn(async () => [
      folder({ Id: 'v', DisplayName: 'Visible', IsHidden: false }),
      folder({ Id: 'h', DisplayName: 'Hidden', IsHidden: true }),
      // IsHidden undefined → treated as visible.
      folder({ Id: 'u', DisplayName: 'Unknown' }),
    ]);
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const rows = await runListFolders(deps, {});
    const ids = rows.map((r) => r.Id);
    expect(ids).toContain('v');
    expect(ids).toContain('u');
    expect(ids).not.toContain('h');
  });

  it('(8b) --include-hidden on surfaces folders whose IsHidden === true', async () => {
    const listFolders = vi.fn(async () => [
      folder({ Id: 'v', DisplayName: 'Visible', IsHidden: false }),
      folder({ Id: 'h', DisplayName: 'Hidden', IsHidden: true }),
    ]);
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const rows = await runListFolders(deps, { includeHidden: true });
    expect(rows.map((r) => r.Id).sort()).toEqual(['h', 'v']);
  });

  it('(9) UpstreamError from the client propagates (mapped by mapHttpError, same taxonomy)', async () => {
    const upstream = new UpstreamError({
      code: 'UPSTREAM_HTTP_503',
      message: 'service unavailable',
      httpStatus: 503,
    });
    const listFolders = vi.fn(async () => {
      throw upstream;
    });
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    let caught: unknown = null;
    try {
      await runListFolders(deps, {});
    } catch (err) {
      caught = err;
    }
    // mapHttpError only translates HTTP-layer errors (http/errors.ts shapes).
    // A pre-wrapped UpstreamError passes through unchanged.
    expect(caught).toBe(upstream);
  });

  it('(10) path segments containing `/` or `\\` are escaped in the emitted Path field', async () => {
    // Backslash is escaped to `\\`, forward-slash is escaped to `\/`.
    const listFolders = vi.fn(async () => [
      folder({ Id: '1', DisplayName: 'a/b' }),
      folder({ Id: '2', DisplayName: 'c\\d' }),
    ]);
    const { deps } = buildDeps({
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const rows = await runListFolders(deps, {});
    expect(rows[0].Path).toBe('a\\/b');
    expect(rows[1].Path).toBe('c\\\\d');
  });
});

describe('list-folders: LIST_FOLDERS_COLUMNS', () => {
  it('(11) declares the §10.7 column order (Path | Unread | Total | Children | Id)', () => {
    const headers = LIST_FOLDERS_COLUMNS.map((c) => c.header);
    expect(headers).toEqual(['Path', 'Unread', 'Total', 'Children', 'Id']);
  });

  it('(12) Id column has no maxWidth (IDs must stay intact per §2.14)', () => {
    const idCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Id');
    expect(idCol).toBeDefined();
    expect(idCol?.maxWidth).toBeUndefined();
  });

  it('(13) extract() helpers coerce missing counters to empty string and fall back to DisplayName when Path is absent', () => {
    // Path ?? DisplayName uses nullish coalescing — only undefined/null fall
    // through (an empty-string Path wins). Pass `Path: undefined` explicitly.
    const row = {
      Id: 'row-1',
      DisplayName: 'Hello',
      Path: undefined,
    } as unknown as Parameters<(typeof LIST_FOLDERS_COLUMNS)[number]['extract']>[0];

    const pathCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Path');
    const unreadCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Unread');
    const totalCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Total');
    const childrenCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Children');
    const idCol = LIST_FOLDERS_COLUMNS.find((c) => c.header === 'Id');

    expect(pathCol?.extract(row)).toBe('Hello');
    expect(unreadCol?.extract(row)).toBe('');
    expect(totalCol?.extract(row)).toBe('');
    expect(childrenCol?.extract(row)).toBe('');
    expect(idCol?.extract(row)).toBe('row-1');
  });
});

// Marker to keep the filename and imports locally symmetric with the
// find-folder spec — also documents which option object is under test.
const _unused: ListFoldersOptions | null = null;
void _unused;
