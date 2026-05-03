// src/commands/list-folders.ts
//
// Enumerate mail folders at a given parent (optionally recursive).
// See project-design.md §10.7 and refined-request-folders.md §5.1.
//
// Wave-4 P5a sole-writer file. Mirrors the pattern of `list-mail.ts` /
// `list-calendar.ts`: config-resolution → ensureSession → createClient →
// resolve parent via the canonical resolver → client.listFolders →
// optional recursive DFS → return typed result. Error mapping is handled by
// the shared `mapHttpError` helper re-exported from `list-mail.ts`.

import type { CliConfig } from '../config/config';
import { UpstreamError } from '../config/errors';
import { parseFolderSpec, resolveFolder } from '../folders/resolver';
import { DEFAULT_LIST_FOLDERS_TOP, MAX_FOLDERS_VISITED } from '../folders/types';
import type { OutlookClient } from '../http/outlook-client';
import type { FolderSummary } from '../http/types';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

export interface ListFoldersDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface ListFoldersOptions {
  /**
   * Well-known alias, display-name path, or `id:<raw>` reference.
   * Default (when omitted / blank): `MsgFolderRoot` (ADR-15).
   */
  parent?: string;
  /**
   * When true, emit the full sub-tree under the resolved parent (breadth-first
   * walk bounded by MAX_FOLDERS_VISITED). Default: false (direct children only).
   */
  recursive?: boolean;
  /**
   * When true, include folders whose `IsHidden === true`. Default: false.
   */
  includeHidden?: boolean;
  /**
   * Upper bound for `$top` passed to each `/childfolders` page.
   * Default: DEFAULT_LIST_FOLDERS_TOP (100). Range: 1..250.
   */
  top?: number;
  /**
   * Forwarded to `resolveFolder` when resolving `--parent`. On ambiguity,
   * pick the oldest candidate (CreatedDateTime asc, Id asc) instead of
   * raising `UsageError('FOLDER_AMBIGUOUS')`. Default: false.
   */
  firstMatch?: boolean;
}

/** Row shape emitted by `run()` — wire `FolderSummary` plus an always-populated
 *  `Path` (escaped) and, in `--recursive` mode, a `Depth` field (0 = direct
 *  child of the resolved parent). */
export interface ListFoldersRow extends FolderSummary {
  Path: string;
  Depth?: number;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Upper bound accepted on `--top` (matches `DEFAULT_LIST_TOP` in folders/types). */
const MAX_TOP = 250;

/** Sentinel value recognised by the resolver for the mailbox root. */
const ROOT_ALIAS = 'MsgFolderRoot';

// ---------------------------------------------------------------------------
// run()
// ---------------------------------------------------------------------------

export async function run(
  deps: ListFoldersDeps,
  opts: ListFoldersOptions = {},
): Promise<ListFoldersRow[]> {
  const top = resolveTop(opts.top);
  const recursive = opts.recursive === true;
  const includeHidden = opts.includeHidden === true;
  const firstMatch = opts.firstMatch === true;

  // Resolve parent. The default (`MsgFolderRoot`) short-circuits without a
  // REST hop — `client.listFolders` accepts the alias verbatim in the URL
  // path (v2.0 contract).
  const parentInput =
    typeof opts.parent === 'string' && opts.parent.length > 0 ? opts.parent : ROOT_ALIAS;

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  let parentId: string;
  if (parentInput === ROOT_ALIAS) {
    parentId = ROOT_ALIAS;
  } else {
    const spec = parseFolderSpec(parentInput);
    try {
      const resolved = await resolveFolder(client, spec, { firstMatch });
      parentId = resolved.Id;
    } catch (err) {
      throw mapHttpError(err);
    }
  }

  try {
    if (!recursive) {
      return await listDirectChildren(client, parentId, top, includeHidden);
    }
    return await listRecursive(client, parentId, top, includeHidden);
  } catch (err) {
    throw mapHttpError(err);
  }
}

// ---------------------------------------------------------------------------
// Non-recursive path
// ---------------------------------------------------------------------------

async function listDirectChildren(
  client: OutlookClient,
  parentId: string,
  top: number,
  includeHidden: boolean,
): Promise<ListFoldersRow[]> {
  const children = await client.listFolders(parentId, top);
  const kept = includeHidden ? children : children.filter(isNotHidden);
  return kept.map((f) => ({
    ...f,
    Path: escapeSegment(f.DisplayName ?? ''),
  }));
}

// ---------------------------------------------------------------------------
// Recursive (breadth-first) walk
// ---------------------------------------------------------------------------

interface Frame {
  id: string;
  /** Escaped display-name path from the resolved anchor down to this folder.
   *  Empty string when the frame represents the anchor itself. */
  path: string;
  depth: number;
}

async function listRecursive(
  client: OutlookClient,
  rootParentId: string,
  top: number,
  includeHidden: boolean,
): Promise<ListFoldersRow[]> {
  const out: ListFoldersRow[] = [];
  const queue: Frame[] = [{ id: rootParentId, path: '', depth: -1 }];

  while (queue.length > 0) {
    const frame = queue.shift() as Frame;
    const children = await client.listFolders(frame.id, top);
    const kept = includeHidden ? children : children.filter(isNotHidden);

    for (const child of kept) {
      if (out.length >= MAX_FOLDERS_VISITED) {
        throw new UpstreamError({
          code: 'UPSTREAM_PAGINATION_LIMIT',
          message:
            `Exceeded ${MAX_FOLDERS_VISITED}-folder cap while walking the folder ` +
            `tree. Narrow --parent or drop --recursive to stay under the cap.`,
        });
      }
      const childPath = composePath(frame.path, child.DisplayName ?? '');
      const childDepth = frame.depth + 1;
      out.push({ ...child, Path: childPath, Depth: childDepth });

      // Descend only when the upstream metadata signals there are children.
      // `ChildFolderCount` may be undefined on some tenants — in that case we
      // still descend, falling back to a GET that may return an empty page.
      const hasChildren =
        typeof child.ChildFolderCount === 'number' ? child.ChildFolderCount > 0 : true;
      if (hasChildren) {
        queue.push({ id: child.Id, path: childPath, depth: childDepth });
      }
    }
  }

  return out;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function resolveTop(raw: number | undefined): number {
  if (raw === undefined) return DEFAULT_LIST_FOLDERS_TOP;
  if (!Number.isInteger(raw) || raw < 1 || raw > MAX_TOP) {
    throw new UsageError(
      `list-folders: --top must be an integer between 1 and ${MAX_TOP} ` + `(got ${String(raw)})`,
    );
  }
  return raw;
}

function isNotHidden(f: FolderSummary): boolean {
  return f.IsHidden !== true;
}

/**
 * Escape a DisplayName segment per the path grammar in project-design §10.5:
 *   - `\`  → `\\`
 *   - `/`  → `\/`
 * The two replacements must run in this order so the backslash-double-escape
 * does not also consume the backslash inserted by the slash rule.
 */
function escapeSegment(segment: string): string {
  return segment.replace(/\\/g, '\\\\').replace(/\//g, '\\/');
}

function composePath(parentPath: string, displayName: string): string {
  const escaped = escapeSegment(displayName);
  return parentPath.length === 0 ? escaped : `${parentPath}/${escaped}`;
}

// ---------------------------------------------------------------------------
// Table rendering helpers (consumed by cli.ts in P6 via a ColumnSpec export)
// ---------------------------------------------------------------------------
//
// Per §10.11, P5a must NOT touch `src/output/formatter.ts` or `src/cli.ts`.
// The command therefore exports a `ColumnSpec`-shaped descriptor that cli.ts
// will pick up during P6 wiring. The extract functions are local so the file
// owns its presentation logic without cross-file coupling.

export const LIST_FOLDERS_COLUMNS: ReadonlyArray<{
  header: string;
  extract: (row: ListFoldersRow) => string;
  maxWidth?: number;
}> = [
  {
    header: 'Path',
    extract: (r) => r.Path ?? r.DisplayName ?? '',
    maxWidth: 48,
  },
  {
    header: 'Unread',
    extract: (r) => (typeof r.UnreadItemCount === 'number' ? String(r.UnreadItemCount) : ''),
  },
  {
    header: 'Total',
    extract: (r) => (typeof r.TotalItemCount === 'number' ? String(r.TotalItemCount) : ''),
  },
  {
    header: 'Children',
    extract: (r) => (typeof r.ChildFolderCount === 'number' ? String(r.ChildFolderCount) : ''),
  },
  {
    header: 'Id',
    extract: (r) => r.Id ?? '',
    // No maxWidth: folder IDs must stay intact for copy-paste into
    // `--parent id:...`, `--folder-id`, or `find-folder id:...`.
  },
];
