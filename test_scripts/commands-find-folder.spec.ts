// test_scripts/commands-find-folder.spec.ts
//
// Command-level tests for `src/commands/find-folder.ts` — `run(deps, spec, opts)`.
// No real HTTP; OutlookClient is injected as a Partial<OutlookClient> cast.
//
// Sources of truth:
//   - src/commands/find-folder.ts (FindFolderDeps, FindFolderOptions, run)
//   - docs/design/project-design.md §10.7 (flag semantics for find-folder)
//   - docs/design/project-design.md §10.5 (path grammar + anchor behaviour)

import { describe, expect, it, vi } from 'vitest';

import type { CliConfig } from '../src/config/config';
import { UsageError } from '../src/commands/list-mail';
import { run as runFindFolder, type FindFolderDeps } from '../src/commands/find-folder';
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
  deps: FindFolderDeps;
  loadSession: ReturnType<typeof vi.fn>;
  saveSession: ReturnType<typeof vi.fn>;
  doAuthCapture: ReturnType<typeof vi.fn>;
  createClient: ReturnType<typeof vi.fn>;
  client: Partial<OutlookClient>;
}

function buildDeps(clientOverrides: Partial<OutlookClient> = {}): BuiltDeps {
  const client: Partial<OutlookClient> = { ...clientOverrides };
  const loadSession = vi.fn(async () => buildFakeSession());
  const saveSession = vi.fn(async () => undefined);
  const doAuthCapture = vi.fn(async () => buildFakeSession());
  const createClient = vi.fn(() => client as OutlookClient);

  const deps: FindFolderDeps = {
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

describe('find-folder: run()', () => {
  it('(1) `id:<raw>` positional calls client.getFolder with the raw id and tags ResolvedVia="id"', async () => {
    const rawId = 'AAMkAGI-opaque-id-xyz';
    const getFolder = vi.fn(async (arg: string) => folder({ Id: arg, DisplayName: 'MyFolder' }));
    const listFolders = vi.fn(async () => {
      throw new Error('listFolders must not be called for id:<raw> spec');
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, `id:${rawId}`);

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith(rawId);
    expect(listFolders).not.toHaveBeenCalled();
    expect(resolved.Id).toBe(rawId);
    expect(resolved.ResolvedVia).toBe('id');
  });

  it('(2) well-known alias positional (`Inbox`) calls client.getFolder("Inbox") with ResolvedVia="wellknown"', async () => {
    const getFolder = vi.fn(async (_arg: string) =>
      folder({ Id: 'inbox-real-id', DisplayName: 'Inbox', WellKnownName: 'inbox' }),
    );
    const listFolders = vi.fn(async () => {
      throw new Error('listFolders must not be called for a wellKnown spec');
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, 'Inbox');

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('Inbox');
    expect(listFolders).not.toHaveBeenCalled();
    expect(resolved.Id).toBe('inbox-real-id');
    expect(resolved.DisplayName).toBe('Inbox');
    expect(resolved.ResolvedVia).toBe('wellknown');
    expect(resolved.Path).toBe('Inbox');
  });

  it('(3) path positional walks the resolver (getFolder for Inbox shortcut + listFolders for each segment)', async () => {
    // "Inbox/Projects" — Inbox is the well-known root shortcut; Projects
    // requires a listFolders pass.
    const inboxId = 'inbox-id-xyz';
    const projectsId = 'projects-id-abc';

    const getFolder = vi.fn(async (arg: string) => {
      if (arg === 'Inbox') {
        return folder({ Id: inboxId, DisplayName: 'Inbox', WellKnownName: 'inbox' });
      }
      throw new Error(`unexpected getFolder(${arg})`);
    });

    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === inboxId) {
        return [folder({ Id: projectsId, DisplayName: 'Projects' })];
      }
      throw new Error(`unexpected listFolders(${parentId})`);
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, 'Inbox/Projects');

    expect(getFolder).toHaveBeenCalledWith('Inbox');
    expect(listFolders).toHaveBeenCalledWith(inboxId);
    expect(resolved.Id).toBe(projectsId);
    expect(resolved.DisplayName).toBe('Projects');
    expect(resolved.ResolvedVia).toBe('path');
    expect(resolved.Path).toBe('Inbox/Projects');
  });

  it('(4a) --anchor on a path spec changes the starting anchor of the walk', async () => {
    // Positional = "Alpha" (a path), --anchor = "id:foo-id" → the resolver
    // must call getFolder("foo-id") for the anchor instead of starting from
    // MsgFolderRoot, then listFolders(foo-id) to find Alpha.
    const anchorId = 'foo-id';
    const alphaId = 'alpha-id';

    const getFolder = vi.fn(async (arg: string) => {
      if (arg === anchorId) {
        return folder({ Id: anchorId, DisplayName: 'AnchorFolder' });
      }
      throw new Error(`unexpected getFolder(${arg})`);
    });

    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === anchorId) {
        return [folder({ Id: alphaId, DisplayName: 'Alpha' })];
      }
      throw new Error(`unexpected listFolders(${parentId})`);
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, 'Alpha', {
      anchor: `id:${anchorId}`,
    });

    expect(getFolder).toHaveBeenCalledWith(anchorId);
    expect(listFolders).toHaveBeenCalledWith(anchorId);
    expect(resolved.Id).toBe(alphaId);
    expect(resolved.ResolvedVia).toBe('path');
  });

  it('(4b) --anchor on a well-known positional is ignored (single getFolder on the alias itself)', async () => {
    // Positional = "Inbox" (wellKnown); --anchor must be silently ignored —
    // the command resolves via a single getFolder("Inbox") call without
    // touching anything related to the anchor.
    const getFolder = vi.fn(async (arg: string) => {
      if (arg === 'Inbox') {
        return folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      }
      throw new Error(`anchor leaked into resolver: getFolder(${arg})`);
    });
    const listFolders = vi.fn(async () => {
      throw new Error('listFolders must not be called for a well-known spec');
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, 'Inbox', {
      anchor: 'SentItems',
    });

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('Inbox');
    expect(listFolders).not.toHaveBeenCalled();
    expect(resolved.ResolvedVia).toBe('wellknown');
  });

  it('(4c) --anchor on an id: positional is ignored (single getFolder on the raw id)', async () => {
    const rawId = 'AAMkAGI-raw';
    const getFolder = vi.fn(async (arg: string) => {
      if (arg === rawId) {
        return folder({ Id: rawId, DisplayName: 'Named' });
      }
      throw new Error(`anchor leaked into resolver: getFolder(${arg})`);
    });
    const listFolders = vi.fn(async () => {
      throw new Error('listFolders must not be called for an id spec');
    });

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, `id:${rawId}`, {
      anchor: 'Inbox',
    });

    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith(rawId);
    expect(listFolders).not.toHaveBeenCalled();
    expect(resolved.ResolvedVia).toBe('id');
  });

  it('(5) --first-match picks the oldest ambiguous sibling (CreatedDateTime asc tiebreaker)', async () => {
    // Two "Alpha" folders under Inbox. Without --first-match: UsageError.
    // With --first-match: the one with the earlier CreatedDateTime wins.
    const inboxId = 'inbox-id';
    const getFolder = vi.fn(async (arg: string) => {
      if (arg === 'Inbox') {
        return folder({ Id: inboxId, DisplayName: 'Inbox', WellKnownName: 'inbox' });
      }
      throw new Error(`unexpected getFolder(${arg})`);
    });
    const listFolders = vi.fn(async () => [
      folder({
        Id: 'alpha-newer',
        DisplayName: 'Alpha',
        CreatedDateTime: '2026-02-01T00:00:00Z',
      }),
      folder({
        Id: 'alpha-older',
        DisplayName: 'Alpha',
        CreatedDateTime: '2026-01-01T00:00:00Z',
      }),
    ]);

    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    const resolved = await runFindFolder(deps, 'Inbox/Alpha', {
      firstMatch: true,
    });

    expect(resolved.Id).toBe('alpha-older');
    expect(resolved.ResolvedVia).toBe('path');
  });

  it('(6) empty positional raises UsageError before any REST call', async () => {
    const getFolder = vi.fn();
    const listFolders = vi.fn();
    const { deps, createClient } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    await expect(runFindFolder(deps, '')).rejects.toBeInstanceOf(UsageError);
    expect(getFolder).not.toHaveBeenCalled();
    expect(listFolders).not.toHaveBeenCalled();
    // Session / client should not even be constructed on a pure-argv failure.
    expect(createClient).not.toHaveBeenCalled();
  });

  it('(7) ambiguity without --first-match propagates as UsageError whose message mentions FOLDER_AMBIGUOUS', async () => {
    const inboxId = 'inbox-id';
    const getFolder = vi.fn(async () =>
      folder({ Id: inboxId, DisplayName: 'Inbox', WellKnownName: 'inbox' }),
    );
    const listFolders = vi.fn(async () => [
      folder({ Id: 'a1', DisplayName: 'Alpha', CreatedDateTime: '2026-01-01T00:00:00Z' }),
      folder({ Id: 'a2', DisplayName: 'Alpha', CreatedDateTime: '2026-02-01T00:00:00Z' }),
    ]);
    const { deps } = buildDeps({
      getFolder: getFolder as OutlookClient['getFolder'],
      listFolders: listFolders as OutlookClient['listFolders'],
    });

    let caught: unknown = null;
    try {
      await runFindFolder(deps, 'Inbox/Alpha');
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(UsageError);
    expect((caught as UsageError).exitCode).toBe(2);
    expect((caught as UsageError).message).toContain('FOLDER_AMBIGUOUS');
  });
});
