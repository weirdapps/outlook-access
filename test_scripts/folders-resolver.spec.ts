// test_scripts/folders-resolver.spec.ts
//
// Unit tests for src/folders/resolver.ts —
// `parseFolderSpec`, `normalizeSegment`, `resolveFolder`, `ensurePath`.
//
// Normative sources:
//   - docs/design/project-design.md §10.5 (Path-resolution algorithm)
//   - docs/design/refined-request-folders.md §11 (acceptance criteria)
//
// No live network. OutlookClient is mocked via `Partial<OutlookClient>` casts.

import { beforeEach, describe, expect, it, vi } from 'vitest';

import {
  ensurePath,
  normalizeSegment,
  parseFolderSpec,
  resolveFolder,
} from '../src/folders/resolver';
import { UsageError } from '../src/commands/list-mail';
import { UpstreamError } from '../src/config/errors';
import { CollisionError } from '../src/http/errors';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary } from '../src/http/types';
import { MAX_PATH_SEGMENTS } from '../src/folders/types';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Build a FolderSummary with sensible defaults. Tests override only the
 * fields they care about.
 */
function folder(
  partial: Partial<FolderSummary> & { Id: string; DisplayName: string },
): FolderSummary {
  return {
    ChildFolderCount: 0,
    UnreadItemCount: 0,
    TotalItemCount: 0,
    IsHidden: false,
    ...partial,
  };
}

/**
 * Build a fake OutlookClient. Only the three folder methods consumed by
 * resolver + ensurePath are stubbed; everything else is left undefined.
 */
function makeFakeClient(overrides: Partial<OutlookClient> = {}): OutlookClient {
  const fallback = vi.fn(async () => {
    throw new Error('fake client: method not stubbed');
  });
  const fake: Partial<OutlookClient> = {
    getFolder: overrides.getFolder ?? (fallback as unknown as OutlookClient['getFolder']),
    listFolders: overrides.listFolders ?? (fallback as unknown as OutlookClient['listFolders']),
    createFolder: overrides.createFolder ?? (fallback as unknown as OutlookClient['createFolder']),
  };
  return fake as OutlookClient;
}

// ---------------------------------------------------------------------------
// parseFolderSpec
// ---------------------------------------------------------------------------

describe('parseFolderSpec', () => {
  it('(1) returns {kind:"id"} for "id:AAMk..."', () => {
    const spec = parseFolderSpec('id:AAMkAGI0PAY9xyz');
    expect(spec).toEqual({ kind: 'id', value: 'AAMkAGI0PAY9xyz' });
  });

  it('(2) returns {kind:"wellKnown"} for case-insensitive alias match when no "/"', () => {
    // Exact canonical form.
    expect(parseFolderSpec('Inbox')).toEqual({
      kind: 'wellKnown',
      value: 'Inbox',
    });
    // Case-insensitive — lowercased input matches canonical PascalCase alias.
    expect(parseFolderSpec('inbox')).toEqual({
      kind: 'wellKnown',
      value: 'Inbox',
    });
    expect(parseFolderSpec('SENTITEMS')).toEqual({
      kind: 'wellKnown',
      value: 'SentItems',
    });
    expect(parseFolderSpec('MsgFolderRoot')).toEqual({
      kind: 'wellKnown',
      value: 'MsgFolderRoot',
    });
  });

  it('(3) returns {kind:"path"} for a plain name that is NOT a well-known alias (single segment)', () => {
    const spec = parseFolderSpec('Projects');
    expect(spec).toEqual({ kind: 'path', value: 'Projects' });
  });

  it('(4) returns {kind:"path"} for a well-known alias when the input contains a separator', () => {
    const spec = parseFolderSpec('Inbox/Projects');
    expect(spec).toEqual({ kind: 'path', value: 'Inbox/Projects' });
  });

  it('(5) escape grammar: `a\\/b` parses to a single segment `a/b`', () => {
    // parseFolderSpec returns a path spec — actual unescaping is proved
    // via resolveFolder below. Here we assert the spec shape is preserved
    // verbatim so the path parser can later unescape correctly.
    const spec = parseFolderSpec('a\\/b');
    expect(spec.kind).toBe('path');
    expect((spec as { value: string }).value).toBe('a\\/b');
  });

  it('(6) escape grammar: `a\\\\b` parses to a single segment `a\\b`', () => {
    const spec = parseFolderSpec('a\\\\b');
    expect(spec.kind).toBe('path');
    // The escape token "\\\\" (two backslashes in the source string) is passed
    // through to the path parser, which will emit a single `\` at compare time.
    expect((spec as { value: string }).value).toBe('a\\\\b');
  });

  it('(7) throws UsageError on empty input', () => {
    expect(() => parseFolderSpec('')).toThrow(UsageError);
  });

  it('(8) throws UsageError on "id:" prefix with empty id', () => {
    expect(() => parseFolderSpec('id:')).toThrow(UsageError);
  });

  // ---------------------------------------------------------------------
  // The following tests exercise parseFolderPath via resolveFolder, because
  // parseFolderSpec defers path-level validation (escapes, trailing slashes,
  // depth cap) to the resolver's internal parser. We don't need network —
  // the resolver fails fast before any client call.
  // ---------------------------------------------------------------------

  it('(9) dangling "\\" at end of a path → UsageError (FOLDER_PATH_INVALID) from resolver', async () => {
    const client = makeFakeClient();
    await expect(resolveFolder(client, { kind: 'path', value: 'foo\\' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('(10) unknown "\\x" escape in a path → UsageError (FOLDER_PATH_INVALID) from resolver', async () => {
    const client = makeFakeClient();
    await expect(resolveFolder(client, { kind: 'path', value: 'foo\\x' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('(11) leading "/" in a path → UsageError (FOLDER_PATH_INVALID)', async () => {
    const client = makeFakeClient();
    await expect(
      resolveFolder(client, { kind: 'path', value: '/Projects' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('(12) trailing "/" in a path → UsageError (FOLDER_PATH_INVALID)', async () => {
    const client = makeFakeClient();
    await expect(
      resolveFolder(client, { kind: 'path', value: 'Projects/' }),
    ).rejects.toBeInstanceOf(UsageError);
  });

  it('(13) "//" in a path → UsageError (FOLDER_PATH_INVALID)', async () => {
    const client = makeFakeClient();
    await expect(resolveFolder(client, { kind: 'path', value: 'a//b' })).rejects.toBeInstanceOf(
      UsageError,
    );
  });

  it('(14) exceeding MAX_PATH_SEGMENTS (17 segments > 16 cap) → UsageError (FOLDER_PATH_INVALID)', async () => {
    const client = makeFakeClient();
    // Build a path with 17 segments.
    const segCount = MAX_PATH_SEGMENTS + 1;
    const longPath = Array.from({ length: segCount }, (_v, i) => `s${i}`).join('/');
    await expect(resolveFolder(client, { kind: 'path', value: longPath })).rejects.toBeInstanceOf(
      UsageError,
    );
  });
});

// ---------------------------------------------------------------------------
// normalizeSegment
// ---------------------------------------------------------------------------

describe('normalizeSegment', () => {
  it('(1) lowercases ASCII', () => {
    expect(normalizeSegment('INBOX')).toBe('inbox');
    expect(normalizeSegment('Projects')).toBe('projects');
  });

  it('(2) NFC-normalizes (decomposed input == composed input after normalize)', () => {
    // "é" as a single composed code point (U+00E9).
    const composed = 'café';
    // "é" as "e" + combining acute (U+0065 U+0301) — canonically equivalent.
    const decomposed = 'café';
    expect(composed).not.toBe(decomposed);
    expect(normalizeSegment(composed)).toBe(normalizeSegment(decomposed));
    expect(normalizeSegment(decomposed)).toBe('café');
  });

  it('(3) does NOT strip leading/trailing whitespace (resolver compares verbatim)', () => {
    // §10.5.3 is explicit: "No trimming, no whitespace collapsing".
    expect(normalizeSegment(' inbox ')).toBe(' inbox ');
    expect(normalizeSegment('a b')).toBe('a b');
  });
});

// ---------------------------------------------------------------------------
// resolveFolder
// ---------------------------------------------------------------------------

describe('resolveFolder', () => {
  describe('{kind:"id"}', () => {
    it('(1) calls getFolder(id) once and tags ResolvedVia:"id"', async () => {
      const target = folder({
        Id: 'AAMkid-1',
        DisplayName: 'Whatever',
        ParentFolderId: 'parent-x',
      });
      const getFolder = vi.fn(async () => target);
      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
      });

      const resolved = await resolveFolder(client, {
        kind: 'id',
        value: 'AAMkid-1',
      });

      expect(getFolder).toHaveBeenCalledTimes(1);
      expect(getFolder).toHaveBeenCalledWith('AAMkid-1');
      expect(resolved.Id).toBe('AAMkid-1');
      expect(resolved.DisplayName).toBe('Whatever');
      expect(resolved.ResolvedVia).toBe('id');
      expect(resolved.Path).toBe('Whatever');
    });
  });

  describe('{kind:"wellKnown"}', () => {
    it('(2) calls getFolder(alias) once and tags ResolvedVia:"wellknown"', async () => {
      const target = folder({ Id: 'wk-id-inbox', DisplayName: 'Inbox' });
      const getFolder = vi.fn(async () => target);
      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
      });

      const resolved = await resolveFolder(client, {
        kind: 'wellKnown',
        value: 'Inbox',
      });

      expect(getFolder).toHaveBeenCalledTimes(1);
      expect(getFolder).toHaveBeenCalledWith('Inbox');
      expect(resolved.ResolvedVia).toBe('wellknown');
      expect(resolved.Path).toBe('Inbox');
      expect(resolved.Id).toBe('wk-id-inbox');
    });
  });

  describe('{kind:"path"} walks', () => {
    it('(3) well-known-wins-at-root: path starting with "Inbox" uses getFolder(alias) for seg 0, then listFolders', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      const projectsFolder = folder({
        Id: 'projects-id',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
      });
      const alphaFolder = folder({
        Id: 'alpha-id',
        DisplayName: 'Alpha',
        ParentFolderId: 'projects-id',
      });

      const getFolder = vi.fn(async (idOrAlias: string) => {
        if (idOrAlias === 'Inbox') return inboxFolder;
        throw new Error(`unexpected getFolder(${idOrAlias})`);
      });
      const listFolders = vi.fn(async (parentId: string) => {
        if (parentId === 'inbox-id') return [projectsFolder];
        if (parentId === 'projects-id') return [alphaFolder];
        throw new Error(`unexpected listFolders(${parentId})`);
      });

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      const resolved = await resolveFolder(client, {
        kind: 'path',
        value: 'Inbox/Projects/Alpha',
      });

      // Segment 0 via getFolder(alias), never via a listFolders hop.
      expect(getFolder).toHaveBeenCalledTimes(1);
      expect(getFolder).toHaveBeenCalledWith('Inbox');
      // Exactly two listFolders hops — one per child segment.
      expect(listFolders).toHaveBeenCalledTimes(2);
      expect(listFolders).toHaveBeenNthCalledWith(1, 'inbox-id');
      expect(listFolders).toHaveBeenNthCalledWith(2, 'projects-id');

      expect(resolved.Id).toBe('alpha-id');
      expect(resolved.DisplayName).toBe('Alpha');
      expect(resolved.ResolvedVia).toBe('path');
      // Path starts with the alias (no escape applied), then joins segments.
      expect(resolved.Path).toBe('Inbox/Projects/Alpha');
    });

    it('(4) non-alias first segment: starts from getFolder("msgfolderroot") and walks via listFolders', async () => {
      const rootFolder = folder({
        Id: 'root-id',
        DisplayName: 'Top of Information Store',
      });
      const topFolder = folder({
        Id: 'top-id',
        DisplayName: 'Projects',
        ParentFolderId: 'root-id',
      });
      const alphaFolder = folder({
        Id: 'alpha-id',
        DisplayName: 'Alpha',
        ParentFolderId: 'top-id',
      });

      const getFolder = vi.fn(async (idOrAlias: string) => {
        if (idOrAlias === 'msgfolderroot') return rootFolder;
        throw new Error(`unexpected getFolder(${idOrAlias})`);
      });
      const listFolders = vi.fn(async (parentId: string) => {
        if (parentId === 'root-id') return [topFolder];
        if (parentId === 'top-id') return [alphaFolder];
        throw new Error(`unexpected listFolders(${parentId})`);
      });

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      const resolved = await resolveFolder(client, {
        kind: 'path',
        value: 'Projects/Alpha',
      });

      expect(getFolder).toHaveBeenCalledTimes(1);
      expect(getFolder).toHaveBeenCalledWith('msgfolderroot');
      expect(listFolders).toHaveBeenCalledTimes(2);
      expect(listFolders).toHaveBeenNthCalledWith(1, 'root-id');
      expect(listFolders).toHaveBeenNthCalledWith(2, 'top-id');

      expect(resolved.Id).toBe('alpha-id');
      expect(resolved.ResolvedVia).toBe('path');
      // No materialized prefix from MsgFolderRoot (§10.5).
      expect(resolved.Path).toBe('Projects/Alpha');
    });

    it('(5) ambiguity at a segment throws UsageError(FOLDER_AMBIGUOUS) by default', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      const twin1 = folder({
        Id: 'twin-1',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2025-01-01T00:00:00Z',
      });
      const twin2 = folder({
        Id: 'twin-2',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2024-01-01T00:00:00Z',
      });

      const getFolder = vi.fn(async () => inboxFolder);
      const listFolders = vi.fn(async () => [twin1, twin2]);

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      await expect(
        resolveFolder(client, { kind: 'path', value: 'Inbox/Projects' }),
      ).rejects.toSatisfy((err: unknown) => {
        return err instanceof UsageError && /FOLDER_AMBIGUOUS/.test(String((err as Error).message));
      });
    });

    it('(6) ambiguity with firstMatch:true picks CreatedDateTime asc then Id asc (ADR-14)', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      // twin1 is newer; twin2 is older → twin2 wins on CreatedDateTime asc.
      const twin1 = folder({
        Id: 'twin-aaa',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2025-06-01T00:00:00Z',
      });
      const twin2 = folder({
        Id: 'twin-zzz',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2024-06-01T00:00:00Z',
      });

      const getFolder = vi.fn(async () => inboxFolder);
      const listFolders = vi.fn(async () => [twin1, twin2]);

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      const resolved = await resolveFolder(
        client,
        { kind: 'path', value: 'Inbox/Projects' },
        { firstMatch: true },
      );
      expect(resolved.Id).toBe('twin-zzz'); // older CreatedDateTime wins.
    });

    it('(7) firstMatch:true — with identical CreatedDateTime, Id asc tiebreaker wins', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      const twin1 = folder({
        Id: 'twin-bbb',
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2024-06-01T00:00:00Z',
      });
      const twin2 = folder({
        Id: 'twin-aaa', // lexicographically smaller Id.
        DisplayName: 'Projects',
        ParentFolderId: 'inbox-id',
        CreatedDateTime: '2024-06-01T00:00:00Z',
      });

      const getFolder = vi.fn(async () => inboxFolder);
      const listFolders = vi.fn(async () => [twin1, twin2]);

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      const resolved = await resolveFolder(
        client,
        { kind: 'path', value: 'Inbox/Projects' },
        { firstMatch: true },
      );
      expect(resolved.Id).toBe('twin-aaa');
    });

    it('(8) 0 matches at a segment throws UpstreamError(UPSTREAM_FOLDER_NOT_FOUND)', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      const getFolder = vi.fn(async () => inboxFolder);
      const listFolders = vi.fn(async () => [
        folder({
          Id: 'other',
          DisplayName: 'Something Else',
          ParentFolderId: 'inbox-id',
        }),
      ]);
      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      await expect(
        resolveFolder(client, { kind: 'path', value: 'Inbox/Projects' }),
      ).rejects.toSatisfy((err: unknown) => {
        return (
          err instanceof UpstreamError &&
          (err as UpstreamError).code === 'UPSTREAM_FOLDER_NOT_FOUND'
        );
      });
    });

    it('(9) enforces MAX_PATH_SEGMENTS cap (already covered in parseFolderSpec, but from path kind too)', async () => {
      const client = makeFakeClient();
      const longPath = Array.from({ length: MAX_PATH_SEGMENTS + 1 }, (_v, i) => `s${i}`).join('/');
      await expect(resolveFolder(client, { kind: 'path', value: longPath })).rejects.toBeInstanceOf(
        UsageError,
      );
    });

    it('(10) case-insensitive matching across decomposed vs composed Unicode', async () => {
      const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
      // Folder on the wire uses decomposed "café" (e + combining acute) in
      // uppercase: "CAFÉ" decomposed → "CAFÉ". Input uses composed
      // lowercase. NFC + case-fold should match them.
      const cafeFolder = folder({
        Id: 'cafe-id',
        DisplayName: 'CAFÉ',
        ParentFolderId: 'inbox-id',
      });

      const getFolder = vi.fn(async () => inboxFolder);
      const listFolders = vi.fn(async () => [cafeFolder]);

      const client = makeFakeClient({
        getFolder: getFolder as unknown as OutlookClient['getFolder'],
        listFolders: listFolders as unknown as OutlookClient['listFolders'],
      });

      const resolved = await resolveFolder(client, {
        kind: 'path',
        value: 'Inbox/café', // composed "café", lowercase.
      });
      expect(resolved.Id).toBe('cafe-id');
    });
  });
});

// ---------------------------------------------------------------------------
// ensurePath
// ---------------------------------------------------------------------------

describe('ensurePath', () => {
  // Each test re-builds its own client. No shared mock state across tests.
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('(1) walk with all segments present → no POST, returns leaf', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });
    const a = folder({
      Id: 'a-id',
      DisplayName: 'A',
      ParentFolderId: 'root-id',
    });
    const b = folder({
      Id: 'b-id',
      DisplayName: 'B',
      ParentFolderId: 'a-id',
    });

    const getFolder = vi.fn(async (idOrAlias: string) => {
      if (idOrAlias === 'MsgFolderRoot') return root;
      throw new Error(`unexpected getFolder(${idOrAlias})`);
    });
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [a];
      if (parentId === 'a-id') return [b];
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder must NOT be called when everything exists');
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    const resolved = await ensurePath(client, ['A', 'B'], {
      createParents: false,
      idempotent: true, // leaf pre-exists, idempotent skips POST.
    });

    expect(resolved.Id).toBe('b-id');
    expect(resolved.ResolvedVia).toBe('path');
    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(2) createParents:false + missing intermediate → throws UsageError(FOLDER_MISSING_PARENT)', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });

    const getFolder = vi.fn(async () => root);
    // Intermediate "A" is MISSING at root-id level.
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [];
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async () => {
      throw new Error(
        'createFolder must NOT be called — parent is missing and createParents is false',
      );
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    await expect(
      ensurePath(client, ['A', 'B'], {
        createParents: false,
        idempotent: false,
      }),
    ).rejects.toSatisfy((err: unknown) => {
      return (
        err instanceof UsageError && /FOLDER_MISSING_PARENT/.test(String((err as Error).message))
      );
    });
    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(3) createParents:true + missing intermediate → createFolder called for it, walk continues', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });
    const aCreated = folder({
      Id: 'a-id',
      DisplayName: 'A',
      ParentFolderId: 'root-id',
    });
    const bCreated = folder({
      Id: 'b-id',
      DisplayName: 'B',
      ParentFolderId: 'a-id',
    });

    const getFolder = vi.fn(async () => root);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return []; // "A" missing.
      if (parentId === 'a-id') return []; // "B" missing too.
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async (parentId: string, displayName: string) => {
      if (parentId === 'root-id' && displayName === 'A') return aCreated;
      if (parentId === 'a-id' && displayName === 'B') return bCreated;
      throw new Error(`unexpected createFolder(${parentId}, ${displayName})`);
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    const resolved = await ensurePath(client, ['A', 'B'], {
      createParents: true,
      idempotent: false,
    });

    expect(resolved.Id).toBe('b-id');
    expect(createFolder).toHaveBeenCalledTimes(2);
    expect(createFolder).toHaveBeenNthCalledWith(1, 'root-id', 'A');
    expect(createFolder).toHaveBeenNthCalledWith(2, 'a-id', 'B');
  });

  it('(4) leaf missing + idempotent:false → creates leaf', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });
    const leaf = folder({
      Id: 'leaf-id',
      DisplayName: 'NewLeaf',
      ParentFolderId: 'root-id',
    });

    const getFolder = vi.fn(async () => root);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return []; // leaf missing.
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async (parentId: string, displayName: string) => {
      if (parentId === 'root-id' && displayName === 'NewLeaf') return leaf;
      throw new Error(`unexpected createFolder(${parentId}, ${displayName})`);
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    const resolved = await ensurePath(client, ['NewLeaf'], {
      createParents: false,
      idempotent: false,
    });

    expect(resolved.Id).toBe('leaf-id');
    expect(createFolder).toHaveBeenCalledTimes(1);
    expect(createFolder).toHaveBeenCalledWith('root-id', 'NewLeaf');
  });

  it('(5) leaf pre-exists + idempotent:false → throws CollisionError', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });
    const existing = folder({
      Id: 'existing-id',
      DisplayName: 'Dup',
      ParentFolderId: 'root-id',
    });

    const getFolder = vi.fn(async () => root);
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'root-id') return [existing];
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder must NOT be called — collision must be raised first');
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    await expect(
      ensurePath(client, ['Dup'], {
        createParents: false,
        idempotent: false,
      }),
    ).rejects.toSatisfy((err: unknown) => {
      return (
        err instanceof CollisionError && (err as CollisionError).code === 'FOLDER_ALREADY_EXISTS'
      );
    });
    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(6) leaf pre-exists + idempotent:true → returns existing leaf, no error, no POST', async () => {
    const root = folder({ Id: 'root-id', DisplayName: 'MsgFolderRoot' });
    const existing = folder({
      Id: 'existing-id',
      DisplayName: 'Dup',
      ParentFolderId: 'root-id',
    });

    const getFolder = vi.fn(async () => root);
    const listFolders = vi.fn(async () => [existing]);
    const createFolder = vi.fn(async () => {
      throw new Error('createFolder must NOT be called — idempotent path is silent');
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    const resolved = await ensurePath(client, ['Dup'], {
      createParents: false,
      idempotent: true,
    });

    expect(resolved.Id).toBe('existing-id');
    expect(resolved.DisplayName).toBe('Dup');
    expect(resolved.ResolvedVia).toBe('path');
    expect(createFolder).not.toHaveBeenCalled();
  });

  it('(7) anchor parameter: walk under --parent Inbox does NOT reach msgfolderroot', async () => {
    const inboxFolder = folder({ Id: 'inbox-id', DisplayName: 'Inbox' });
    const projects = folder({
      Id: 'projects-id',
      DisplayName: 'Projects',
      ParentFolderId: 'inbox-id',
    });

    const getFolder = vi.fn(async (idOrAlias: string) => {
      if (idOrAlias === 'Inbox') return inboxFolder;
      // An attempt to hit msgfolderroot here is the failure mode we guard against.
      throw new Error(`ensurePath with --parent Inbox must not call getFolder("${idOrAlias}")`);
    });
    const listFolders = vi.fn(async (parentId: string) => {
      if (parentId === 'inbox-id') return []; // Projects missing → create it.
      throw new Error(`unexpected listFolders(${parentId})`);
    });
    const createFolder = vi.fn(async (parentId: string, displayName: string) => {
      if (parentId === 'inbox-id' && displayName === 'Projects') return projects;
      throw new Error(`unexpected createFolder(${parentId}, ${displayName})`);
    });

    const client = makeFakeClient({
      getFolder: getFolder as unknown as OutlookClient['getFolder'],
      listFolders: listFolders as unknown as OutlookClient['listFolders'],
      createFolder: createFolder as unknown as OutlookClient['createFolder'],
    });

    const resolved = await ensurePath(client, ['Projects'], {
      createParents: false,
      idempotent: false,
      anchor: { kind: 'wellKnown', value: 'Inbox' },
    });

    expect(resolved.Id).toBe('projects-id');
    // Only getFolder('Inbox') ever ran — never 'msgfolderroot'.
    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('Inbox');
    const calledAliases = getFolder.mock.calls.map((c) => c[0]);
    expect(calledAliases).not.toContain('msgfolderroot');
    expect(calledAliases).not.toContain('MsgFolderRoot');
    expect(createFolder).toHaveBeenCalledWith('inbox-id', 'Projects');
  });
});
