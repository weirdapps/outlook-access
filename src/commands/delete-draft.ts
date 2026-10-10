// src/commands/delete-draft.ts
//
// Delete one or more DRAFT messages. Nothing else: every id is fetched first
// and the server's own `IsDraft` flag decides. A sent or received message is
// refused before any DELETE is issued, so this cannot be used to clear mail.
//
// Exchange moves a deleted item to Deleted Items, so a mistake is recoverable
// from Outlook. The loop is single-threaded, like move-mail.

import type { CliConfig } from '../config/config';
import { UpstreamError } from '../config/errors';
import type { OutlookClient } from '../http/outlook-client';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface DeleteDraftDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface DeleteDraftOptions {
  /** If true, per-message failures are collected into `failed[]` instead of short-circuiting. */
  continueOnError?: boolean;
}

export interface DeleteDraftResult {
  deleted: Array<{ id: string; subject: string }>;
  failed: Array<{ id: string; error: { code: string; httpStatus?: number; message?: string } }>;
  summary: { requested: number; deleted: number; failed: number };
}

export async function run(
  deps: DeleteDraftDeps,
  messageIds: string[],
  opts: DeleteDraftOptions = {},
): Promise<DeleteDraftResult> {
  if (!Array.isArray(messageIds) || messageIds.length === 0) {
    throw new UsageError('delete-draft: at least one <messageId> positional argument is required');
  }
  for (const id of messageIds) {
    if (typeof id !== 'string' || id.length === 0) {
      throw new UsageError(
        'delete-draft: <messageId> positional arguments must be non-empty strings',
      );
    }
  }
  const continueOnError = opts.continueOnError === true;

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const deleted: DeleteDraftResult['deleted'] = [];
  const failed: DeleteDraftResult['failed'] = [];

  for (const id of messageIds) {
    try {
      const m = await client.getMessage(id, { select: ['Id', 'Subject', 'IsDraft'] });
      if (m?.IsDraft !== true) {
        throw new UpstreamError({
          code: 'NOT_A_DRAFT',
          message: `message '${id}' is not a draft; delete-draft only deletes unsent drafts.`,
        });
      }
      await client.deleteMessage(id);
      deleted.push({ id, subject: m.Subject ?? '' });
    } catch (err) {
      const mapped = err instanceof UpstreamError ? err : mapHttpError(err);
      if (!continueOnError) throw mapped;
      failed.push({ id, error: toError(mapped) });
    }
  }

  return {
    deleted,
    failed,
    summary: { requested: messageIds.length, deleted: deleted.length, failed: failed.length },
  };
}

function toError(err: unknown): { code: string; httpStatus?: number; message?: string } {
  if (err instanceof UpstreamError) {
    return { code: err.code, httpStatus: err.httpStatus, message: err.message };
  }
  const maybe = err as { code?: unknown; message?: unknown };
  return {
    code: typeof maybe.code === 'string' && maybe.code.length > 0 ? maybe.code : 'UPSTREAM_UNKNOWN',
    message: typeof maybe.message === 'string' ? maybe.message : String(err),
  };
}
