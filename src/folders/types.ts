/**
 * CLI-layer types for the folder-management feature.
 *
 * Wire-shaped REST types (`FolderSummary`, `FolderCreateRequest`,
 * `MoveMessageRequest`) live in `src/http/types.ts` next to the other REST
 * shapes. This file owns the *CLI* shapes — the tagged-union `FolderSpec`
 * consumed by the resolver, the resolved form surfaced by `find-folder`, and
 * the result shapes surfaced by `create-folder` / `move-mail`.
 *
 * Normative source: project-design.md §10.3.2 (types) + §10.11 (publish list).
 */

import type { FolderSummary } from '../http/types';

// ---------------------------------------------------------------------------
// Well-known aliases
// ---------------------------------------------------------------------------

/**
 * Exhaustive PascalCase alias list accepted in the v2.0 URL path.
 * Source: refined-request-folders.md §6.2 +
 *         outlook-v2-folder-pagination-filter.md §References #3.
 */
export type WellKnownAlias =
  | 'Inbox'
  | 'SentItems'
  | 'Drafts'
  | 'DeletedItems'
  | 'Archive'
  | 'JunkEmail'
  | 'Outbox'
  | 'MsgFolderRoot'
  | 'RecoverableItemsDeletions';

/**
 * Canonical, frozen list of every well-known alias the v2.0 URL path accepts.
 * Order matches the precedence used by `matchesWellKnownAlias` (P4).
 */
export const WELL_KNOWN_ALIASES: readonly WellKnownAlias[] = Object.freeze([
  'Inbox',
  'SentItems',
  'Drafts',
  'DeletedItems',
  'Archive',
  'JunkEmail',
  'Outbox',
  'MsgFolderRoot',
  'RecoverableItemsDeletions',
]);

// ---------------------------------------------------------------------------
// FolderSpec — the canonical folder reference understood by the resolver
// ---------------------------------------------------------------------------

/**
 * The canonical folder reference understood by the resolver. Discriminated
 * tagged union — every caller uses exactly one kind and the resolver branches
 * on `kind` without heuristics.
 *
 *   { kind: 'wellKnown', value: 'Inbox' }                    → alias in URL path, no lookup
 *   { kind: 'id',        value: 'AAMkAGI...' }               → raw opaque id
 *   { kind: 'path',      value: 'Projects/Alpha', parent? }  → segmented walk under an anchor
 *
 * The optional `parent` on a path spec is the anchor for a bare/non-absolute
 * path (default: `{ kind: 'wellKnown', value: 'MsgFolderRoot' }` — ADR-15).
 */
export type FolderSpec =
  | { kind: 'wellKnown'; value: WellKnownAlias }
  | { kind: 'id'; value: string }
  | { kind: 'path'; value: string; parent?: FolderSpec };

// ---------------------------------------------------------------------------
// ResolvedFolder — output of the resolver
// ---------------------------------------------------------------------------

/**
 * Result of `resolveFolder`. Extends the wire `FolderSummary` with a
 * materialized `Path` and a `ResolvedVia` provenance tag that is surfaced
 * verbatim in the JSON output of `find-folder`.
 *
 * `Id` and `DisplayName` are promoted to mandatory on the resolved form —
 * the resolver guarantees both are populated.
 */
export interface ResolvedFolder extends FolderSummary {
  /** Always present (mandatory on the resolved form). */
  Id: string;
  DisplayName: string;
  /** Materialized path from the anchor down, using the escape grammar in §10.5. */
  Path: string;
  /** How the resolver arrived at this folder. Serialised in `find-folder` JSON. */
  ResolvedVia: 'wellknown' | 'path' | 'id';
}

// ---------------------------------------------------------------------------
// CreateFolderResult — output of `create-folder`
// ---------------------------------------------------------------------------

/** One entry in `CreateFolderResult.created[]`. */
export interface CreateFolderSegment {
  Id: string;
  DisplayName: string;
  /** Slash-delimited path from the anchor down to (and including) this segment. */
  Path: string;
  ParentFolderId: string;
  /**
   * True when the segment already existed (only visible under
   * `--idempotent` or `--create-parents`).
   */
  PreExisting: boolean;
}

export interface CreateFolderResult {
  /** One entry per processed segment (existing and newly created alike). */
  created: CreateFolderSegment[];
  /** Convenience pointer — equals `created[created.length - 1]`. */
  leaf: CreateFolderSegment;
  /**
   * True iff every leaf resolution path was pre-existing (no POST issued
   * for the leaf).
   */
  idempotent: boolean;
}

// ---------------------------------------------------------------------------
// MoveMailResult — output of `move-mail`
// ---------------------------------------------------------------------------

/** One entry in `MoveMailResult.moved[]` (success path). */
export interface MoveEntry {
  sourceId: string;
  newId: string;
}

/** One entry in `MoveMailResult.failed[]` (populated only under `--continue-on-error`). */
export interface MoveFailedEntry {
  sourceId: string;
  error: { code: string; httpStatus?: number; message?: string };
}

/** Destination block included in `MoveMailResult` for id ↔ path symmetry. */
export interface MoveDestination {
  Id: string;
  Path: string;
  DisplayName: string;
}

export interface MoveMailResult {
  destination: MoveDestination;
  moved: MoveEntry[];
  failed: MoveFailedEntry[];
  summary: { requested: number; moved: number; failed: number };
}

// ---------------------------------------------------------------------------
// Safety-cap constants
// ---------------------------------------------------------------------------

/**
 * Maximum number of segments accepted by `parseFolderPath`. Paths longer
 * than this raise `UsageError('FOLDER_PATH_INVALID', ...)`.
 * Pinned in plan-002 §P1.
 */
export const MAX_PATH_SEGMENTS = 16;

/**
 * Maximum number of `@odata.nextLink` pages followed by `listAll<T>` for a
 * single collection before raising `ApiError('PAGINATION_LIMIT', ...)`.
 * Pinned in plan-002 §P1 / research §6.
 */
export const MAX_FOLDER_PAGES = 50;

/**
 * Whole-tree cap for `list-folders --recursive`. Raises
 * `UpstreamError('UPSTREAM_PAGINATION_LIMIT', ...)` when exceeded.
 * Pinned in plan-002 §P1.
 */
export const MAX_FOLDERS_VISITED = 5000;

/** Default `$top` used by `listAll<T>` on the first page. */
export const DEFAULT_LIST_TOP = 250;

/** Default `$top` surfaced by `list-folders` (CLI option `--top`, range 1..250). */
export const DEFAULT_LIST_FOLDERS_TOP = 100;
