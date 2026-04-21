// src/commands/list-mail.ts
//
// List recent messages in a well-known folder.
// See project-design.md §2.13.3 and refined spec §5.3.

import type { CliConfig } from '../config/config';
import {
  AuthError as CliAuthError,
  OutlookCliError,
  UpstreamError,
} from '../config/errors';
import type { OutlookClient } from '../http/outlook-client';
import {
  ApiError,
  AuthError as HttpAuthError,
  NetworkError,
} from '../http/errors';
import type { MessageSummary, ODataListResponse } from '../http/types';
import type { SessionFile } from '../session/schema';
import { isExpired } from '../session/store';
import { parseFolderSpec, resolveFolder } from '../folders/resolver';

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
}

export const ALLOWED_FOLDERS = [
  'Inbox',
  'SentItems',
  'Drafts',
  'DeletedItems',
  'Archive',
] as const;

const DEFAULT_SELECT =
  'Id,Subject,From,ReceivedDateTime,HasAttachments,IsRead,WebLink';

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
): Promise<MessageSummary[]> {
  // Resolve effective option values (fall back to CliConfig defaults).
  const top = typeof opts.top === 'number' ? opts.top : deps.config.listMailTop;
  if (!Number.isInteger(top) || top < 1 || top > 100) {
    throw new UsageError(
      `list-mail: --top must be an integer between 1 and 100 (got ${String(top)})`,
    );
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
  const hasFolderId =
    typeof opts.folderId === 'string' && opts.folderId.length > 0;
  const hasFolderParent =
    typeof opts.folderParent === 'string' && opts.folderParent.length > 0;

  if (hasFolder && hasFolderId) {
    throw new UsageError(
      'list-mail: --folder and --folder-id are mutually exclusive.',
    );
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
    typeof opts.select === 'string' && opts.select.length > 0
      ? opts.select
      : DEFAULT_SELECT;

  // Session load. If missing/expired and auto-reauth is allowed, capture a
  // fresh session before we build the client.
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  // Path A: --folder-id → use the id verbatim (no resolver hop).
  if (hasFolderId) {
    try {
      return await client.listMessagesInFolder(opts.folderId as string, {
        top,
        select: select.split(',').map((s) => s.trim()).filter((s) => s.length > 0),
        orderBy: 'ReceivedDateTime desc',
      });
    } catch (err) {
      throw mapHttpError(err);
    }
  }

  // Effective --folder value (falls back to CliConfig.listMailFolder = "Inbox").
  const folder = hasFolder ? (opts.folder as string) : deps.config.listMailFolder;

  // Path B (fast path): the value is one of the original five well-known
  // aliases (Inbox, SentItems, Drafts, DeletedItems, Archive). No resolver
  // hop — the alias is sent verbatim in the URL path, preserving the exact
  // request shape used before the folder feature landed.
  const isFastPathAlias = (ALLOWED_FOLDERS as readonly string[]).includes(folder);
  if (isFastPathAlias && !hasFolderParent) {
    const restPath = `/api/v2.0/me/MailFolders/${encodeURIComponent(folder)}/messages`;
    const query = {
      $top: String(top),
      $orderby: 'ReceivedDateTime desc',
      $select: select,
    };

    try {
      const resp = await client.get<ODataListResponse<MessageSummary>>(
        restPath,
        query,
      );
      return Array.isArray(resp.value) ? resp.value : [];
    } catch (err) {
      throw mapHttpError(err);
    }
  }

  // Path C (resolver): the value is a path (Inbox/Projects/...), a non-
  // fast-path well-known alias (JunkEmail / Outbox / MsgFolderRoot /
  // RecoverableItemsDeletions), or a bare name anchored by --folder-parent.
  let resolvedId: string;
  try {
    const spec = parseFolderSpec(folder);
    // Attach the anchor only when meaningful (path form).
    const finalSpec =
      spec.kind === 'path' && hasFolderParent
        ? { ...spec, parent: parseFolderSpec(opts.folderParent as string) }
        : spec;
    const resolved = await resolveFolder(client, finalSpec);
    resolvedId = resolved.Id;
  } catch (err) {
    throw mapHttpError(err);
  }

  try {
    return await client.listMessagesInFolder(resolvedId, {
      top,
      select: select.split(',').map((s) => s.trim()).filter((s) => s.length > 0),
      orderBy: 'ReceivedDateTime desc',
    });
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
    const cliCode =
      err.reason === 'NO_AUTO_REAUTH'
        ? 'AUTH_NO_REAUTH'
        : 'AUTH_401_AFTER_RETRY';
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
