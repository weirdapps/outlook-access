// src/commands/list-mail.ts
//
// List recent messages in a well-known folder.
// See project-design.md §2.13.3 and refined spec §5.3.

import type { CliConfig } from '../config/config';
import { AuthError as CliAuthError, OutlookCliError, UpstreamError } from '../config/errors';
import type { OutlookClient } from '../http/outlook-client';
import { ApiError, AuthError as HttpAuthError, NetworkError } from '../http/errors';
import { buildReceivedDateFilter, FilterError } from '../http/filter-builder';
import type { MessageSummary } from '../http/types';
import type { SessionFile } from '../session/schema';
import { isExpired } from '../session/store';
import { parseFolderSpec, resolveFolder } from '../folders/resolver';
import { parseTimestamp } from '../util/dates';

export interface ListMailDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface ListMailOptions {
  top?: number;
  folder?: string;
  select?: string;
  folderId?: string;
  folderParent?: string;
  /** ISO-8601 UTC timestamp; include only messages with ReceivedDateTime >= this. */
  since?: string;
  /** ISO-8601 UTC timestamp; include only messages with ReceivedDateTime < this. */
  until?: string;
  /**
   * Lower bound (inclusive, `ge`) on ReceivedDateTime. Accepts ISO-8601 or
   * the keyword grammar `now` / `now + Nd` / `now - Nd`. Mutually exclusive
   * with `--since`.
   */
  from?: string;
  /**
   * Upper bound (exclusive, `lt`) on ReceivedDateTime. Same grammar as
   * `--from`. Mutually exclusive with `--until`.
   */
  to?: string;
  /** When true, walk @odata.nextLink up to `max` results (default 10000). */
  all?: boolean;
  /** Hard cap on total results when `all` is true. Defaults to 10000. */
  max?: number;
  /**
   * When true, return just the count of matching messages (server-side via
   * `$count=true`) instead of the messages themselves. Ignores `--top` and
   * `--select`. Works alongside every folder flag and the date window.
   */
  justCount?: boolean;
}

/** Result shape returned by `run()` when `justCount` is true. */
export interface ListMailCountResult {
  count: number;
  exact: boolean;
}

/** Default safety cap when --all is set without --max. */
export const DEFAULT_MAX_RESULTS = 10000;
/** Absolute hard ceiling — refuses to walk more than this many results in one call. */
export const ABSOLUTE_MAX_RESULTS = 100_000;

export const ALLOWED_FOLDERS = ['Inbox', 'SentItems', 'Drafts', 'DeletedItems', 'Archive'] as const;

const DEFAULT_SELECT = 'Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead,WebLink';

/**
 * Raised for user-input validation failures that commander does not catch.
 * Maps to exit code 2 via the top-level handler in cli.ts.
 */
export class UsageError extends OutlookCliError {
  public readonly code: string = 'BAD_USAGE';
  public readonly exitCode: number = 2;
}

export async function run(
  deps: ListMailDeps,
  opts: ListMailOptions = {},
): Promise<MessageSummary[] | ListMailCountResult> {
  const justCount = opts.justCount === true;

  // Resolve effective option values (fall back to CliConfig defaults).
  // Skip --top validation in count mode — the flag is ignored there.
  const top = typeof opts.top === 'number' ? opts.top : deps.config.listMailTop;
  if (!justCount && (!Number.isInteger(top) || top < 1 || top > 1000)) {
    throw new UsageError(
      `list-mail: --top must be an integer between 1 and 1000 (got ${String(top)})`,
    );
  }

  // --just-count is incompatible with --all (count is one HTTP call by design).
  if (justCount && opts.all === true) {
    throw new UsageError(
      'list-mail: --just-count and --all are mutually exclusive ' +
        '(count uses server-side $count=true and returns in one request).',
    );
  }

  // Pagination flags (--all / --max). Validate eagerly.
  const fetchAll = opts.all === true;
  const maxResults = opts.max ?? DEFAULT_MAX_RESULTS;
  if (!Number.isInteger(maxResults) || maxResults < 1) {
    throw new UsageError(`list-mail: --max must be a positive integer (got ${String(maxResults)})`);
  }
  if (maxResults > ABSOLUTE_MAX_RESULTS) {
    throw new UsageError(
      `list-mail: --max cannot exceed ${ABSOLUTE_MAX_RESULTS} (got ${String(maxResults)})`,
    );
  }

  // Mutually exclusive: --since/--until (legacy) vs --from/--to (v1.2.0+).
  const hasSinceUntil =
    (typeof opts.since === 'string' && opts.since.length > 0) ||
    (typeof opts.until === 'string' && opts.until.length > 0);
  const hasFromTo =
    (typeof opts.from === 'string' && opts.from.length > 0) ||
    (typeof opts.to === 'string' && opts.to.length > 0);
  if (hasSinceUntil && hasFromTo) {
    throw new UsageError(
      'list-mail: --since/--until and --from/--to are mutually exclusive ' +
        '(prefer --from/--to which accepts the now/now±Nd keyword grammar).',
    );
  }

  // Filter clause: prefer --from/--to (parseTimestamp), fall back to legacy
  // --since/--until (buildReceivedDateFilter from filter-builder module).
  let filter: string;
  if (hasFromTo) {
    filter = buildFromToFilter(opts.from, opts.to);
  } else {
    try {
      filter = buildReceivedDateFilter(opts.since, opts.until);
    } catch (err) {
      if (err instanceof FilterError) {
        throw new UsageError(`list-mail: ${err.message}`);
      }
      throw err;
    }
  }

  // Additive flags (§10.7):
  //   --folder-id     : XOR with --folder — bypasses the resolver entirely.
  //   --folder-parent : anchor for a path / bare-name in --folder. Meaningful
  //                     only alongside --folder; passing it with --folder-id
  //                     (or alone, without --folder) → UsageError exit 2.
  //   --folder        : widened to accept paths (Inbox/Projects/...) and all
  //                     well-known aliases in addition to the original fast-
  //                     path set (ALLOWED_FOLDERS).
  const hasFolder = typeof opts.folder === 'string' && opts.folder.length > 0;
  const hasFolderId = typeof opts.folderId === 'string' && opts.folderId.length > 0;
  const hasFolderParent = typeof opts.folderParent === 'string' && opts.folderParent.length > 0;

  if (hasFolder && hasFolderId) {
    throw new UsageError('list-mail: --folder and --folder-id are mutually exclusive.');
  }
  if (hasFolderParent && hasFolderId) {
    throw new UsageError(
      'list-mail: --folder-parent cannot be combined with --folder-id ' +
        '(the id is absolute — no anchor is needed).',
    );
  }
  if (hasFolderParent && !hasFolder) {
    throw new UsageError(
      'list-mail: --folder-parent requires --folder (it is an anchor for ' +
        'a bare name / path in --folder).',
    );
  }

  const select =
    typeof opts.select === 'string' && opts.select.length > 0 ? opts.select : DEFAULT_SELECT;

  // Session load. If missing/expired and auto-reauth is allowed, capture a
  // fresh session before we build the client.
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const selectArr = select
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  const listOpts = {
    top,
    select: selectArr,
    orderBy: 'ReceivedDateTime desc',
    filter: filter.length > 0 ? filter : undefined,
  };

  // Resolve the target folderId via one of three paths.
  // Path A: --folder-id → use the id verbatim (no resolver hop).
  // Path B (fast path): well-known alias used verbatim in URL.
  // Path C (resolver): path form, non-fast-path alias, or anchored bare name.
  let targetFolderId: string;
  if (hasFolderId) {
    targetFolderId = opts.folderId as string;
  } else {
    const folder = hasFolder ? (opts.folder as string) : deps.config.listMailFolder;
    const isFastPathAlias = (ALLOWED_FOLDERS as readonly string[]).includes(folder);
    if (isFastPathAlias && !hasFolderParent) {
      // Fast-path alias: pass verbatim. listMessagesInFolder URL-encodes it.
      targetFolderId = folder;
    } else {
      // Path C — resolver
      try {
        const spec = parseFolderSpec(folder);
        const finalSpec =
          spec.kind === 'path' && hasFolderParent
            ? { ...spec, parent: parseFolderSpec(opts.folderParent as string) }
            : spec;
        const resolved = await resolveFolder(client, finalSpec);
        targetFolderId = resolved.Id;
      } catch (err) {
        throw mapHttpError(err);
      }
    }
  }

  // Single dispatch point: count, paginate, or single-page.
  try {
    if (justCount) {
      return await client.countMessagesInFolder(targetFolderId, {
        filter: filter.length > 0 ? filter : undefined,
      });
    }
    if (fetchAll) {
      const result = await client.listMessagesInFolderAll(targetFolderId, listOpts, maxResults);
      if (result.truncated) {
        process.stderr.write(
          JSON.stringify({
            code: 'max_results_reached',
            message: `--max=${maxResults} cap hit; ${result.messages.length} returned, more available`,
            hint: 'increase --max or split query with --since/--until',
          }) + '\n',
        );
      }
      return result.messages;
    }
    return await client.listMessagesInFolder(targetFolderId, listOpts);
  } catch (err) {
    throw mapHttpError(err);
  }
}

// ---------------------------------------------------------------------------
// Shared helpers (re-used by every other command in this unit)
// ---------------------------------------------------------------------------

/**
 * Common session load / refresh path. Returns a SessionFile ready to use.
 * Respects `--no-auto-reauth`.
 */
export async function ensureSession(deps: {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
}): Promise<SessionFile> {
  const cached = await deps.loadSession(deps.sessionPath);

  if (cached !== null && !isExpired(cached)) {
    return cached;
  }
  if (deps.config.noAutoReauth) {
    throw new CliAuthError(
      'AUTH_NO_REAUTH',
      'Session is missing or expired and --no-auto-reauth was set.',
    );
  }
  const fresh = await deps.doAuthCapture();
  await deps.saveSession(deps.sessionPath, fresh);
  return fresh;
}

/**
 * Translate HTTP-layer errors into the CLI's error taxonomy so the top-level
 * handler in cli.ts can map them to exit codes directly.
 */
export function mapHttpError(err: unknown): unknown {
  if (err instanceof HttpAuthError) {
    // Distinguish the two 401 paths per design §2.8:
    //   - noAutoReauth first 401 → AUTH_NO_REAUTH
    //   - after single retry     → AUTH_401_AFTER_RETRY
    const cliCode = err.reason === 'NO_AUTO_REAUTH' ? 'AUTH_NO_REAUTH' : 'AUTH_401_AFTER_RETRY';
    return new CliAuthError(cliCode, err.message, err);
  }
  if (err instanceof ApiError) {
    return new UpstreamError({
      code: `UPSTREAM_HTTP_${err.httpStatus}`,
      message: err.message,
      httpStatus: err.httpStatus,
      requestId: err.requestId,
      url: err.url,
      cause: err,
    });
  }
  if (err instanceof NetworkError) {
    return new UpstreamError({
      code: err.timedOut ? 'UPSTREAM_TIMEOUT' : 'UPSTREAM_NETWORK',
      message: err.message,
      url: err.url,
      cause: err,
    });
  }
  return err;
}

/**
 * Build a `$filter=ReceivedDateTime ge X and ReceivedDateTime lt Y` expression
 * from `--from` / `--to` inputs (each accepting ISO-8601 or the keyword
 * grammar `now` / `now + Nd` / `now - Nd`). Returns `''` when both bounds
 * are unset. Raises `UsageError` when either bound is malformed.
 *
 * Convention: lower bound is INCLUSIVE (`ge`), upper bound is EXCLUSIVE
 * (`lt`), matching the calendar-view convention.
 */
function buildFromToFilter(from: string | undefined, to: string | undefined): string {
  const hasFrom = typeof from === 'string' && from.length > 0;
  const hasTo = typeof to === 'string' && to.length > 0;
  if (!hasFrom && !hasTo) {
    return '';
  }
  const parts: string[] = [];
  if (hasFrom) {
    const r = parseTimestamp(from as string);
    if (!r.ok) {
      throw new UsageError(`list-mail: --from is ${r.reason}`);
    }
    parts.push(`ReceivedDateTime ge ${r.iso}`);
  }
  if (hasTo) {
    const r = parseTimestamp(to as string);
    if (!r.ok) {
      throw new UsageError(`list-mail: --to is ${r.reason}`);
    }
    parts.push(`ReceivedDateTime lt ${r.iso}`);
  }
  return parts.join(' and ');
}
