// src/folders/resolver.ts
//
// Canonical folder resolver. One module owns every piece of path / alias /
// NFC / case-fold / ambiguity / well-known precedence logic consumed by the
// folder-aware commands.
//
// Normative sources:
//   - docs/design/project-design.md §10.4 (OutlookClient API)
//   - docs/design/project-design.md §10.5 (path-resolution algorithm)
//   - docs/design/project-design.md §10.6 (error handling)
//   - docs/design/plan-002-folders.md §P4
//   - docs/research/outlook-v2-folder-pagination-filter.md §§5-6
//
// Escape grammar (normative, mirrored from project-design §10.5):
//   A path string is a slash-separated sequence of display-name segments.
//   Exactly two characters need escaping inside a segment: `/` (encoded as
//   `\/`) and `\` (encoded as `\\`). No other escape sequences exist;
//   whitespace is preserved and Unicode passes through verbatim (subject to
//   NFC normalization at compare time).

import { UpstreamError } from '../config/errors';
import { CollisionError } from '../http/errors';
import { UsageError } from '../commands/list-mail';
import type { OutlookClient } from '../http/outlook-client';
import type { FolderSummary } from '../http/types';
import {
  MAX_PATH_SEGMENTS,
  WELL_KNOWN_ALIASES,
  type FolderSpec,
  type ResolvedFolder,
  type WellKnownAlias,
} from './types';

// ---------------------------------------------------------------------------
// parseFolderSpec — CLI input → FolderSpec tagged union
// ---------------------------------------------------------------------------

/**
 * Parse a CLI-supplied folder reference (as produced by `--parent`, `--to`,
 * `--folder`, or the positional `<query>` of `find-folder`) into the
 * canonical tagged union consumed by `resolveFolder`.
 *
 *   - `"id:AAMk..."`                → { kind: 'id', value: rest }
 *   - `"Inbox"` (case-insensitive exact match against WELL_KNOWN_ALIASES and
 *      no `/` nor `\/` present in the input)
 *                                  → { kind: 'wellKnown', value: canonical }
 *   - anything else (path with `/` separators, optionally escaped per the
 *     grammar in the header comment)
 *                                  → { kind: 'path', value: input }
 *   - empty / blank input           → UsageError('BAD_USAGE') (exit 2)
 *
 * The function does NOT perform the path walk or REST calls; it is a pure
 * string → spec translation. Callers pass the spec to `resolveFolder`.
 */
export function parseFolderSpec(input: string): FolderSpec {
  if (typeof input !== 'string' || input.length === 0) {
    throw new UsageError('folder spec: input is required (got empty string)');
  }

  // Raw-id prefix — consume verbatim. The remainder is an opaque Outlook id.
  if (input.startsWith('id:')) {
    const id = input.slice(3);
    if (id.length === 0) {
      throw new UsageError("folder spec: 'id:' prefix with empty id");
    }
    return { kind: 'id', value: id };
  }

  // Well-known alias — only if the input contains no separator (neither bare
  // `/` nor the `\/` escape). A string that contains a separator is always
  // treated as a path, even if its first segment happens to be a well-known
  // alias. (The "well-known-wins-at-root" rule is applied inside
  // `resolveFolder`, not here — see §10.5.)
  if (!containsPathSeparator(input)) {
    const alias = matchesWellKnownAlias(input);
    if (alias !== null) {
      return { kind: 'wellKnown', value: alias };
    }
  }

  // Everything else is a path. Syntactic validation (empty segments, bad
  // escape, depth cap) is deferred to `parseFolderPath`, which runs inside
  // `resolveFolder` so the resolver retains a single validation surface.
  return { kind: 'path', value: input };
}

// ---------------------------------------------------------------------------
// Path parser + segment normalization
// ---------------------------------------------------------------------------

/**
 * Unicode normalization form + simple case-fold for client-side DisplayName
 * matching. Used by `normalizeSegment` (below) and at compare time inside
 * `resolveFolder`. See §10.5 "Case-folding + NFC normalization".
 */
export function normalizeSegment(s: string): string {
  return s.normalize('NFC').toLocaleLowerCase('en-US');
}

/**
 * Split + unescape a slash-separated folder path into NFC-normalized
 * segments, enforcing the depth cap and escape rules.
 *
 * Grammar (see the header comment for the normative statement):
 *   - `/` inside a DisplayName is encoded as `\/`.
 *   - `\` inside a DisplayName is encoded as `\\`.
 *   - Any other `\<x>` sequence is a syntax error.
 *   - Empty segments (leading `/`, trailing `/`, or `//`) are syntax errors.
 *   - A dangling trailing `\` is a syntax error.
 *
 * Raises `UsageError('FOLDER_PATH_INVALID', ...)` on any violation.
 */
function parseFolderPath(input: string): string[] {
  // NOSONAR S3776 - path parsing with escape handling
  if (typeof input !== 'string' || input.length === 0) {
    throw new UsageError('folder path: empty path (FOLDER_PATH_INVALID)');
  }

  const segments: string[] = [];
  let current = '';
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (c === '\\') {
      if (i + 1 >= input.length) {
        throw new UsageError('folder path: dangling escape at end of input (FOLDER_PATH_INVALID)');
      }
      const next = input[i + 1];
      if (next === '/') {
        current += '/';
        i += 2;
      } else if (next === '\\') {
        current += '\\';
        i += 2;
      } else {
        throw new UsageError(`folder path: unknown escape '\\${next}' (FOLDER_PATH_INVALID)`);
      }
    } else if (c === '/') {
      if (current === '') {
        throw new UsageError('folder path: empty segment (FOLDER_PATH_INVALID)');
      }
      segments.push(current.normalize('NFC'));
      current = '';
      i += 1;
    } else {
      current += c;
      i += 1;
    }
  }

  if (current === '' && segments.length === 0) {
    throw new UsageError('folder path: empty path (FOLDER_PATH_INVALID)');
  }
  if (current === '') {
    // Trailing '/'.
    throw new UsageError('folder path: empty segment (FOLDER_PATH_INVALID)');
  }
  segments.push(current.normalize('NFC'));

  if (segments.length > MAX_PATH_SEGMENTS) {
    throw new UsageError(
      `folder path: depth ${segments.length} exceeds cap ${MAX_PATH_SEGMENTS} ` +
        `(FOLDER_PATH_INVALID)`,
    );
  }

  return segments;
}

/**
 * Does `input` contain an unescaped path separator? Used by `parseFolderSpec`
 * to decide between the well-known fast path and the path walk.
 *
 * Implementation note: a bare `\` not followed by `/` or `\` is NOT a
 * separator here — the definitive decision about its legality is made inside
 * `parseFolderPath` later. We only care whether the caller *intended* a
 * path.
 */
function containsPathSeparator(input: string): boolean {
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (c === '\\') {
      // Skip the escape pair without interpreting it as a separator, even if
      // malformed — `parseFolderPath` will raise later.
      i += 2;
      continue;
    }
    if (c === '/') return true;
    i += 1;
  }
  return false;
}

/**
 * Exact case-insensitive match against WELL_KNOWN_ALIASES. Returns the
 * canonical PascalCase form or null. This is the public helper consumed by
 * `parseFolderSpec` and the "well-known-wins-at-root" shortcut in
 * `resolveFolder`.
 */
function matchesWellKnownAlias(input: string): WellKnownAlias | null {
  if (typeof input !== 'string' || input.length === 0) return null;
  const normalized = normalizeSegment(input);
  for (const alias of WELL_KNOWN_ALIASES) {
    if (normalizeSegment(alias) === normalized) return alias;
  }
  return null;
}

// ---------------------------------------------------------------------------
// resolveFolder — the path-walk workhorse (§10.5)
// ---------------------------------------------------------------------------

interface ResolveOptions {
  /** When true, pick the first match on ambiguity instead of raising. */
  firstMatch?: boolean;
}

/**
 * Resolve a `FolderSpec` to a single `ResolvedFolder` per the algorithm in
 * project-design §10.5.
 *
 *   - `kind: 'id'`        → `client.getFolder(value)`; `ResolvedVia: 'id'`.
 *   - `kind: 'wellKnown'` → `client.getFolder(alias)`;  `ResolvedVia: 'wellknown'`.
 *   - `kind: 'path'`      → walk segment-by-segment, starting from
 *                            `MsgFolderRoot`; `ResolvedVia: 'path'`.
 *
 * Well-known-wins-at-root: when `kind === 'path'` and the first segment
 * (case-insensitive) matches a well-known alias, segment 0 is resolved via
 * `client.getFolder(alias)` instead of a child lookup under MsgFolderRoot
 * (§10.5 "Well-known precedence").
 *
 * Ambiguity at a segment: when `opts.firstMatch !== true`, raises
 * `UsageError('FOLDER_AMBIGUOUS')` (exit 2). Otherwise candidates are sorted
 * by `(CreatedDateTime asc, Id asc)` per ADR-14 and the first is chosen.
 *
 * Zero matches at a segment: raises
 * `UpstreamError('UPSTREAM_FOLDER_NOT_FOUND')` (exit 5).
 */
export async function resolveFolder(
  client: OutlookClient,
  spec: FolderSpec,
  opts: ResolveOptions = {},
): Promise<ResolvedFolder> {
  switch (spec.kind) {
    case 'id': {
      const folder = await client.getFolder(spec.value);
      return toResolved(folder, folder.DisplayName ?? '', 'id');
    }

    case 'wellKnown': {
      const folder = await client.getFolder(spec.value);
      return toResolved(folder, spec.value, 'wellknown');
    }

    case 'path': {
      return await resolvePath(client, spec, opts);
    }

    default: {
      // Exhaustiveness check.
      const _exhaustive: never = spec;
      throw new Error(`resolveFolder: unknown FolderSpec kind (${String(_exhaustive)})`);
    }
  }
}

async function resolvePath( // NOSONAR S3776 - path resolution with retries
  client: OutlookClient,
  spec: Extract<FolderSpec, { kind: 'path' }>,
  opts: ResolveOptions,
): Promise<ResolvedFolder> {
  const segments = parseFolderPath(spec.value);

  // Choose the anchor. Default: MsgFolderRoot (ADR-15).
  const anchorSpec: FolderSpec = spec.parent ?? { kind: 'wellKnown', value: 'MsgFolderRoot' };

  // Well-known-wins-at-root: applies only when the anchor is MsgFolderRoot
  // (the default), and only to segment 0. A shadowed top-level user folder
  // called "Inbox" is reachable via `--parent MsgFolderRoot --first-match`
  // or by id — see §10.5 "Well-known precedence".
  const anchorIsRoot = anchorSpec.kind === 'wellKnown' && anchorSpec.value === 'MsgFolderRoot';

  let currentId: string;
  let currentPath: string;
  let startIndex = 0;
  let lastFolder: FolderSummary | null = null;

  if (anchorIsRoot) {
    const rootAlias = matchesWellKnownAlias(segments[0]);
    if (rootAlias !== null) {
      // Segment 0 resolves via the well-known alias shortcut.
      const folder = await client.getFolder(rootAlias);
      lastFolder = folder;
      currentId = folder.Id;
      currentPath = rootAlias;
      startIndex = 1;
    } else {
      // Start from MsgFolderRoot; no materialized prefix (§10.5).
      const root = await client.getFolder('msgfolderroot');
      currentId = root.Id;
      currentPath = '';
    }
  } else {
    const anchor = await resolveFolder(client, anchorSpec, opts);
    currentId = anchor.Id;
    currentPath = anchor.Path;
  }

  for (let i = startIndex; i < segments.length; i++) {
    const segment = segments[i];
    const children = await client.listFolders(currentId);
    const matches = children.filter(
      (c) =>
        typeof c.DisplayName === 'string' &&
        normalizeSegment(c.DisplayName) === normalizeSegment(segment),
    );

    if (matches.length === 0) {
      throw new UpstreamError({
        code: 'UPSTREAM_FOLDER_NOT_FOUND',
        message: `Folder segment '${segment}' was not found under parent id ` + `'${currentId}'.`,
      });
    }

    let chosen: FolderSummary;
    if (matches.length === 1) {
      chosen = matches[0];
    } else {
      if (opts.firstMatch !== true) {
        const candidateIds = matches.map((m) => m.Id).join(', ');
        throw new UsageError(
          `folder path: segment '${segment}' is ambiguous under parent ` +
            `'${currentId}'; ${matches.length} candidates: [${candidateIds}] ` +
            `(FOLDER_AMBIGUOUS). Re-run with --first-match to pick the oldest, ` +
            `or pass a raw id via 'id:...'.`,
        );
      }
      // ADR-14 tiebreaker: CreatedDateTime asc, then Id asc.
      chosen = [...matches].sort(compareByCreatedAscThenId)[0];
    }

    lastFolder = chosen;
    currentId = chosen.Id;
    currentPath =
      currentPath === '' ? escapeSegment(segment) : `${currentPath}/${escapeSegment(segment)}`;
  }

  if (lastFolder === null) {
    // Can only happen when segments.length === 0 — which parseFolderPath
    // already rejects. Defensive.
    throw new UsageError('folder path: empty path after parse (FOLDER_PATH_INVALID)');
  }

  return toResolved(lastFolder, currentPath, 'path');
}

// ---------------------------------------------------------------------------
// ensurePath — used by `create-folder --create-parents`
// ---------------------------------------------------------------------------

interface EnsurePathOptions {
  /** When true, missing intermediate segments are created; otherwise a
   *  missing non-leaf segment raises UsageError('FOLDER_MISSING_PARENT'). */
  createParents: boolean;
  /** When true, a `FOLDER_ALREADY_EXISTS` collision is swallowed and the
   *  pre-existing folder is returned instead. */
  idempotent: boolean;
  /** Anchor folder spec. Defaults to `{ kind: 'wellKnown', value: 'MsgFolderRoot' }`.
   *  When specified, the walk starts from the resolved anchor instead of the
   *  mailbox root (§10.7 `--parent` contract). */
  anchor?: FolderSpec;
}

/**
 * Walk the given path under `MsgFolderRoot`, creating missing levels when
 * `createParents` is true and treating `CollisionError('FOLDER_ALREADY_EXISTS')`
 * as a no-op when `idempotent` is true.
 *
 * Returns the leaf as a `ResolvedFolder` with `ResolvedVia: 'path'`. Does
 * NOT return the per-segment `PreExisting` list — that is assembled by the
 * `create-folder` command on top of this primitive (which keeps this module
 * free of CLI-shaped result types).
 *
 * Raises:
 *   - UsageError('FOLDER_MISSING_PARENT') when a non-leaf segment is missing
 *     and `createParents` is false.
 *   - CollisionError('FOLDER_ALREADY_EXISTS') when a leaf already exists and
 *     `idempotent` is false (surfaced by OutlookClient.createFolder).
 *   - UsageError('FOLDER_PATH_INVALID') / 'FOLDER_AMBIGUOUS' per resolveFolder
 *     rules (segments are walked via the same client-side matching).
 *   - UpstreamError on any HTTP failure.
 */
export async function ensurePath( // NOSONAR S3776 - recursive path creation
  client: OutlookClient,
  segments: string[],
  opts: EnsurePathOptions,
): Promise<ResolvedFolder> {
  if (!Array.isArray(segments) || segments.length === 0) {
    throw new UsageError('ensurePath: segments must be a non-empty array (FOLDER_PATH_INVALID)');
  }
  if (segments.length > MAX_PATH_SEGMENTS) {
    throw new UsageError(
      `ensurePath: depth ${segments.length} exceeds cap ${MAX_PATH_SEGMENTS} ` +
        `(FOLDER_PATH_INVALID)`,
    );
  }

  // Resolve the anchor. Default: MsgFolderRoot (§10.7). When the caller
  // supplies an explicit `anchor` (e.g. from `--parent Inbox`), the walk
  // starts from that resolved folder. The well-known-wins-at-root rule is
  // NOT applied here — the caller is expected to have rejected "Inbox at
  // root" for the leaf up front (§10.7 validation table).
  const anchorSpec: FolderSpec = opts.anchor ?? { kind: 'wellKnown', value: 'MsgFolderRoot' };
  const anchorResolved = await resolveFolder(client, anchorSpec);
  let currentId = anchorResolved.Id;
  let currentPath = anchorResolved.Path === 'MsgFolderRoot' ? '' : anchorResolved.Path;
  let lastFolder: FolderSummary = anchorResolved;

  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i];
    const isLeaf = i === segments.length - 1;

    // Normalize the segment to NFC for compare; the wire DisplayName is sent
    // verbatim on the POST body, but compares go through normalizeSegment.
    const nfcSegment = segment.normalize('NFC');

    const children = await client.listFolders(currentId);
    const normalizedTarget = normalizeSegment(nfcSegment);
    const matches = children.filter(
      (c) =>
        typeof c.DisplayName === 'string' && normalizeSegment(c.DisplayName) === normalizedTarget,
    );

    if (matches.length > 1) {
      const candidateIds = matches.map((m) => m.Id).join(', ');
      throw new UsageError(
        `ensurePath: segment '${segment}' is ambiguous under parent ` +
          `'${currentId}'; ${matches.length} candidates: [${candidateIds}] ` +
          `(FOLDER_AMBIGUOUS).`,
      );
    }

    let next: FolderSummary;
    if (matches.length === 1) {
      // Segment already exists.
      //   - Intermediate segment: always advance without POST (regardless of
      //     `idempotent`) — a missing-parent error is explicitly about
      //     *missing* parents, not pre-existing ones.
      //   - Leaf segment: without `idempotent`, pre-existence is a collision
      //     and must raise `CollisionError` (exit 6) — symmetric with the
      //     single-name branch and with AC-CREATE-COLLISION. With
      //     `idempotent`, advance and let the caller mark `PreExisting: true`.
      if (isLeaf && !opts.idempotent) {
        throw new CollisionError({
          code: 'FOLDER_ALREADY_EXISTS',
          message: `A folder named '${segment}' already exists under parent ` + `'${currentId}'.`,
          path: segment,
          parentId: currentId,
        });
      }
      next = matches[0];
    } else {
      // Segment missing.
      if (!isLeaf && !opts.createParents) {
        throw new UsageError(
          `ensurePath: missing intermediate folder '${segment}' under parent ` +
            `'${currentId}' (FOLDER_MISSING_PARENT). Re-run with ` +
            `--create-parents to create it.`,
        );
      }
      try {
        next = await client.createFolder(currentId, nfcSegment);
      } catch (err) {
        // Idempotent swallow: the folder may have been created by a concurrent
        // run between our list and our POST. Re-list and pick up the pre-
        // existing child.
        if (opts.idempotent && err instanceof CollisionError) {
          const retry = await client.listFolders(currentId);
          const normTarget = normalizeSegment(nfcSegment);
          const recovered = retry.filter(
            (c) =>
              typeof c.DisplayName === 'string' && normalizeSegment(c.DisplayName) === normTarget,
          );
          if (recovered.length === 0) {
            // Collision reported but folder not visible post-retry. Forward
            // the original CollisionError — the caller will exit 6.
            throw err;
          }
          if (recovered.length > 1) {
            const ids = recovered.map((m) => m.Id).join(', ');
            throw new UsageError(
              `ensurePath: segment '${segment}' is ambiguous after idempotent ` +
                `recovery under parent '${currentId}'; candidates: [${ids}] ` +
                `(FOLDER_AMBIGUOUS).`,
            );
          }
          next = recovered[0];
        } else {
          throw err;
        }
      }
    }

    lastFolder = next;
    currentId = next.Id;
    currentPath =
      currentPath === '' ? escapeSegment(segment) : `${currentPath}/${escapeSegment(segment)}`;
  }

  return toResolved(lastFolder, currentPath, 'path');
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

function toResolved(
  folder: FolderSummary,
  path: string,
  resolvedVia: ResolvedFolder['ResolvedVia'],
): ResolvedFolder {
  // The resolver guarantees Id + DisplayName are populated — FolderSummary
  // types them as required on the wire, so we forward them verbatim.
  return {
    ...folder,
    Id: folder.Id,
    DisplayName: folder.DisplayName,
    Path: path,
    ResolvedVia: resolvedVia,
  };
}

/**
 * ADR-14 tiebreaker: sort candidates by CreatedDateTime ascending, then by
 * Id ascending. Folders missing CreatedDateTime sort LAST (stable bucket)
 * so deterministic-but-unknown metadata does not shadow folders with a
 * known timestamp.
 */
function compareByCreatedAscThenId(a: FolderSummary, b: FolderSummary): number {
  const ca = a.CreatedDateTime;
  const cb = b.CreatedDateTime;
  if (ca !== undefined && cb !== undefined) {
    if (ca < cb) return -1;
    if (ca > cb) return 1;
  } else if (ca !== undefined && cb === undefined) {
    return -1;
  } else if (ca === undefined && cb !== undefined) {
    return 1;
  }
  // Fall through to Id tiebreaker.
  if (a.Id < b.Id) return -1;
  if (a.Id > b.Id) return 1;
  return 0;
}

/**
 * Inverse of the parse grammar — used to render the materialized `Path` on
 * `ResolvedFolder`. Escapes `\` then `/` so the output round-trips through
 * `parseFolderPath` verbatim.
 */
function escapeSegment(segment: string): string {
  return segment.replace(/\\/g, '\\\\').replace(/\//g, '\\/');
}
