// src/commands/create-folder.ts
//
// Create (or idempotently reuse) a mail folder under an anchor.
//
// Normative sources:
//   - docs/design/project-design.md §10.7 (CLI surface contract)
//   - docs/design/project-design.md §10.9 (concurrency / idempotency)
//   - docs/design/refined-request-folders.md §5.3
//   - docs/design/plan-002-folders.md §P5c
//
// Semantics (per §10.7):
//   - Plain name (no '/'): create directly under `--parent` (default
//     `MsgFolderRoot`).
//   - Path (has '/'): split, resolve parent chain; with `--create-parents`,
//     create any missing intermediate segments; always create the leaf.
//   - `--idempotent`: if the leaf (or the whole path) is already present,
//     return it verbatim with `PreExisting: true` instead of failing with
//     exit 6. The race-recovery path is owned by `ensurePath` /
//     `client.createFolder` (see §10.9).
//   - Without `--idempotent`: a collision raises `CollisionError` (exit 6).
//
// This command is the *sole writer* of `src/commands/create-folder.ts`; it
// delegates every REST + path-walk primitive to:
//   - `resolveFolder` (anchor resolution for the single-name case)
//   - `ensurePath`    (full path-walk for the nested case)
//   - `client.createFolder` (terminal REST call for the single-name case)
// and NEVER replicates segment-matching / alias / NFC logic in-file.

import type { CliConfig } from '../config/config';
import { UpstreamError } from '../config/errors';
import { CollisionError } from '../http/errors';
import type { OutlookClient } from '../http/outlook-client';
import type { FolderSummary } from '../http/types';
import type { SessionFile } from '../session/schema';

import {
  ensurePath,
  normalizeSegment,
  parseFolderSpec,
  resolveFolder,
} from '../folders/resolver';
import type {
  CreateFolderResult,
  CreateFolderSegment,
  FolderSpec,
  ResolvedFolder,
} from '../folders/types';
import { WELL_KNOWN_ALIASES } from '../folders/types';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

// ---------------------------------------------------------------------------
// Public deps + options
// ---------------------------------------------------------------------------

export interface CreateFolderDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface CreateFolderOptions {
  /** Anchor spec (well-known alias, path, or `id:...`). Default: `MsgFolderRoot`. */
  parent?: string;
  /** Create missing intermediate segments on a nested path. */
  createParents?: boolean;
  /** Swallow `FOLDER_ALREADY_EXISTS` and return the existing folder instead. */
  idempotent?: boolean;
  /** Override DisplayName used for the leaf segment's POST body. */
  displayName?: string;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export async function run(
  deps: CreateFolderDeps,
  path: string,
  opts: CreateFolderOptions = {},
): Promise<CreateFolderResult> {
  // ----------------------------- argv validation --------------------------
  if (typeof path !== 'string' || path.length === 0) {
    throw new UsageError('create-folder: <path> is required');
  }

  const createParents = opts.createParents === true;
  const idempotent = opts.idempotent === true;

  // Distinguish plain-name vs. multi-segment path. `parseFolderSpec` returns:
  //   - kind: 'id'        when input starts with 'id:'
  //   - kind: 'wellKnown' when input (no '/') matches a well-known alias
  //   - kind: 'path'      otherwise (bare name without '/' still becomes 'path')
  // For create-folder, neither 'id' nor 'wellKnown' is an acceptable target:
  // the user cannot "create" an id, and well-known aliases cannot be created.
  const posSpec = parseFolderSpec(path);
  if (posSpec.kind === 'id') {
    throw new UsageError(
      "create-folder: <path> cannot be a raw id (got 'id:...')",
    );
  }
  if (posSpec.kind === 'wellKnown') {
    throw new UsageError(
      `create-folder: <path> '${path}' is a well-known alias and cannot be ` +
        `created at the mailbox root.`,
    );
  }

  const isNestedPath = containsUnescapedSlash(path);

  // ----------------------------- session + client -------------------------
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  try {
    if (!isNestedPath) {
      return await runSingleName(client, path, opts, createParents, idempotent);
    }
    return await runNestedPath(client, path, opts, createParents, idempotent);
  } catch (err) {
    // CollisionError carries its own exit code (6) and must not be re-wrapped
    // as an UpstreamError (which would exit 5). Let it propagate verbatim.
    if (err instanceof CollisionError) {
      throw err;
    }
    // UsageError is already an OutlookCliError with exit 2. Same handling.
    if (err instanceof UsageError) {
      throw err;
    }
    throw mapHttpError(err);
  }
}

// ---------------------------------------------------------------------------
// Single-name branch — "create directly under --parent"
// ---------------------------------------------------------------------------

async function runSingleName(
  client: OutlookClient,
  leafName: string,
  opts: CreateFolderOptions,
  _createParents: boolean,
  idempotent: boolean,
): Promise<CreateFolderResult> {
  // Resolve the parent anchor. Default: MsgFolderRoot (ADR-15).
  const parentInput =
    typeof opts.parent === 'string' && opts.parent.length > 0
      ? opts.parent
      : 'MsgFolderRoot';
  const parentSpec: FolderSpec = parseFolderSpec(parentInput);
  const parent: ResolvedFolder = await resolveFolder(client, parentSpec);

  // Determine the DisplayName that will hit the wire. The `--display-name`
  // override replaces the last path segment verbatim (refined §5.3).
  const displayName =
    typeof opts.displayName === 'string' && opts.displayName.length > 0
      ? opts.displayName
      : leafName;

  // §10.7 validation: forbid creating a well-known alias at the root.
  if (
    isWellKnownAliasName(displayName) &&
    isMsgFolderRootAnchor(parent, parentSpec)
  ) {
    throw new UsageError(
      `create-folder: cannot create a folder named '${displayName}' at the ` +
        `mailbox root (well-known alias is reserved).`,
    );
  }

  try {
    const created = await client.createFolder(parent.Id, displayName);
    return buildSingleSegmentResult(created, parent, displayName, false);
  } catch (err) {
    if (err instanceof CollisionError && idempotent) {
      // Idempotent recovery: re-list the parent and locate the existing
      // folder by DisplayName. Authoritative because the POST collision
      // confirmed the server-side state.
      const existing = await lookupChildByDisplayName(
        client,
        parent.Id,
        displayName,
      );
      if (existing === null) {
        // Race-window / hidden-folder edge case — the POST reported a
        // collision but the re-list cannot locate the sibling. Surface the
        // original collision verbatim (exit 6) rather than invent state.
        throw err;
      }
      return buildSingleSegmentResult(existing, parent, displayName, true);
    }
    throw err;
  }
}

function buildSingleSegmentResult(
  folder: FolderSummary,
  parent: ResolvedFolder,
  leafDisplayName: string,
  preExisting: boolean,
): CreateFolderResult {
  const parentPath = parent.Path;
  const segmentPath =
    parentPath === ''
      ? escapeSegment(leafDisplayName)
      : `${parentPath}/${escapeSegment(leafDisplayName)}`;

  const segment: CreateFolderSegment = {
    Id: folder.Id,
    DisplayName: folder.DisplayName,
    Path: segmentPath,
    ParentFolderId: folder.ParentFolderId ?? parent.Id,
    PreExisting: preExisting,
  };
  return {
    created: [segment],
    leaf: segment,
    idempotent: preExisting,
  };
}

// ---------------------------------------------------------------------------
// Nested-path branch — delegate the walk to `ensurePath`
// ---------------------------------------------------------------------------

async function runNestedPath(
  client: OutlookClient,
  pathInput: string,
  opts: CreateFolderOptions,
  createParents: boolean,
  idempotent: boolean,
): Promise<CreateFolderResult> {
  // Resolve the anchor (§10.7 `--parent`; default MsgFolderRoot).
  const parentInput =
    typeof opts.parent === 'string' && opts.parent.length > 0
      ? opts.parent
      : 'MsgFolderRoot';
  const anchorSpec: FolderSpec = parseFolderSpec(parentInput);
  const anchorIsRoot =
    anchorSpec.kind === 'wellKnown' && anchorSpec.value === 'MsgFolderRoot';

  // §10.7 validation: reject a path whose LAST segment is a well-known
  // alias ONLY when the anchor is MsgFolderRoot — otherwise `Inbox/X/Inbox`
  // is a legitimate user folder tree. The bare-alias-at-root case was
  // rejected by `parseFolderSpec` at the top of `run`.
  const leafSegment = extractLastSegment(pathInput);
  if (anchorIsRoot && isWellKnownAliasName(leafSegment)) {
    throw new UsageError(
      `create-folder: leaf segment '${leafSegment}' is a well-known alias ` +
        `and cannot be created at the mailbox root.`,
    );
  }

  // Idempotent fast path: try to resolve the full path BEFORE creating any
  // segment. If it already resolves, short-circuit with PreExisting=true.
  // The anchor spec is attached so `tryResolveExistingPath` walks under the
  // same anchor `ensurePath` will use.
  if (idempotent) {
    const existing = await tryResolveExistingPath(client, pathInput, anchorSpec);
    if (existing !== null) {
      const segment: CreateFolderSegment = {
        Id: existing.Id,
        DisplayName:
          typeof opts.displayName === 'string' && opts.displayName.length > 0
            ? opts.displayName
            : existing.DisplayName,
        Path: existing.Path,
        ParentFolderId: existing.ParentFolderId ?? '',
        PreExisting: true,
      };
      return {
        created: [segment],
        leaf: segment,
        idempotent: true,
      };
    }
  }

  // Delegate the full walk (including race-safe idempotent recovery on any
  // segment) to ensurePath. It returns the resolved leaf but not per-segment
  // PreExisting flags; we therefore surface a single-entry `created[]` for
  // the leaf, which is the scriptable handle callers need.
  const segments = splitPath(pathInput);
  const leaf = await ensurePath(client, segments, {
    createParents,
    idempotent,
    anchor: anchorSpec,
  });

  // Honour `--display-name` override on the leaf segment post-walk. The
  // DisplayName on the wire was the path segment; the override affects the
  // CLI payload only if the two differ and the caller explicitly asked.
  const leafDisplayName =
    typeof opts.displayName === 'string' && opts.displayName.length > 0
      ? opts.displayName
      : leaf.DisplayName;

  const segment: CreateFolderSegment = {
    Id: leaf.Id,
    DisplayName: leafDisplayName,
    Path: leaf.Path,
    ParentFolderId: leaf.ParentFolderId ?? '',
    PreExisting: false,
  };
  return {
    created: [segment],
    leaf: segment,
    idempotent: false,
  };
}

// ---------------------------------------------------------------------------
// Small local utilities
// ---------------------------------------------------------------------------

/**
 * Best-effort resolution of an existing path. Returns null when the path does
 * not (yet) exist, otherwise returns the resolved folder. Non-not-found
 * upstream failures re-throw verbatim so the caller can surface them.
 */
async function tryResolveExistingPath(
  client: OutlookClient,
  pathInput: string,
  anchorSpec: FolderSpec,
): Promise<ResolvedFolder | null> {
  const spec: FolderSpec = { kind: 'path', value: pathInput, parent: anchorSpec };
  try {
    return await resolveFolder(client, spec);
  } catch (err) {
    if (
      err instanceof UpstreamError &&
      err.code === 'UPSTREAM_FOLDER_NOT_FOUND'
    ) {
      return null;
    }
    throw err;
  }
}

/**
 * Re-list a parent's direct children and return the (unique) child whose
 * DisplayName matches `target` after NFC + case-fold normalization. Returns
 * null when no child matches; throws UsageError on ambiguity (two children
 * share the same normalized DisplayName — rare, but surfaceable).
 */
async function lookupChildByDisplayName(
  client: OutlookClient,
  parentId: string,
  target: string,
): Promise<FolderSummary | null> {
  const children = await client.listFolders(parentId);
  const normalized = normalizeSegment(target);
  const matches = children.filter(
    (c) =>
      typeof c.DisplayName === 'string' &&
      normalizeSegment(c.DisplayName) === normalized,
  );
  if (matches.length === 0) return null;
  if (matches.length > 1) {
    const ids = matches.map((m) => m.Id).join(', ');
    throw new UsageError(
      `create-folder: folder '${target}' is ambiguous under parent ` +
        `'${parentId}' after idempotent recovery; candidates: [${ids}] ` +
        `(FOLDER_AMBIGUOUS).`,
    );
  }
  return matches[0];
}

/**
 * True iff `name` (after NFC + case-fold) matches one of the canonical
 * well-known aliases (Inbox, SentItems, Drafts, ...). Case-insensitive.
 */
function isWellKnownAliasName(name: string): boolean {
  if (typeof name !== 'string' || name.length === 0) return false;
  const normalized = normalizeSegment(name);
  for (const alias of WELL_KNOWN_ALIASES) {
    if (normalizeSegment(alias) === normalized) return true;
  }
  return false;
}

/**
 * True iff the resolved parent is (or was supplied as) `MsgFolderRoot`. Used
 * for the "cannot create Inbox at root" validation.
 */
function isMsgFolderRootAnchor(
  parent: ResolvedFolder,
  parentSpec: FolderSpec,
): boolean {
  if (parentSpec.kind === 'wellKnown' && parentSpec.value === 'MsgFolderRoot') {
    return true;
  }
  // Defensive: compare against the WellKnownName surfaced by Outlook for the
  // root. `getFolder('msgfolderroot')` returns `WellKnownName === 'msgfolderroot'`.
  if (
    typeof parent.WellKnownName === 'string' &&
    parent.WellKnownName.toLowerCase() === 'msgfolderroot'
  ) {
    return true;
  }
  return false;
}

/**
 * Does `input` contain an unescaped path separator? Mirrors the helper used
 * inside resolver.ts (kept local because it is not exported).
 *
 * `\/` is an escaped literal slash (part of a single DisplayName segment),
 * `\\` is an escaped backslash, any other `\<x>` is deferred to the parser.
 */
function containsUnescapedSlash(input: string): boolean {
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (c === '\\') {
      // Skip the next character regardless; a dangling/invalid escape will
      // be caught by the path parser downstream.
      i += 2;
      continue;
    }
    if (c === '/') return true;
    i += 1;
  }
  return false;
}

/**
 * Split a path into raw segments, mirroring the escape grammar used by
 * `parseFolderPath`. Does NOT enforce the depth cap or the empty-segment
 * check — `ensurePath` / `resolveFolder` own those.
 *
 * Kept local because `parseFolderPath` itself is not exported from the
 * resolver module.
 */
function splitPath(input: string): string[] {
  const segments: string[] = [];
  let current = '';
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (c === '\\') {
      if (i + 1 >= input.length) {
        throw new UsageError(
          'create-folder: dangling escape at end of path (FOLDER_PATH_INVALID)',
        );
      }
      const next = input[i + 1];
      if (next === '/') {
        current += '/';
        i += 2;
      } else if (next === '\\') {
        current += '\\';
        i += 2;
      } else {
        throw new UsageError(
          `create-folder: unknown escape '\\${next}' (FOLDER_PATH_INVALID)`,
        );
      }
    } else if (c === '/') {
      if (current === '') {
        throw new UsageError(
          'create-folder: empty segment (FOLDER_PATH_INVALID)',
        );
      }
      segments.push(current);
      current = '';
      i += 1;
    } else {
      current += c;
      i += 1;
    }
  }
  if (current === '') {
    throw new UsageError(
      'create-folder: empty segment (FOLDER_PATH_INVALID)',
    );
  }
  segments.push(current);
  return segments;
}

/**
 * Return the last display-name segment from a path, respecting the escape
 * grammar. Used for the "leaf is a well-known alias at root" validation.
 */
function extractLastSegment(input: string): string {
  const segs = splitPath(input);
  return segs[segs.length - 1] ?? '';
}

/**
 * Inverse of `splitPath` for a single segment. Escapes `\` first, then `/`,
 * so the output round-trips through `splitPath` verbatim.
 */
function escapeSegment(segment: string): string {
  return segment.replace(/\\/g, '\\\\').replace(/\//g, '\\/');
}
