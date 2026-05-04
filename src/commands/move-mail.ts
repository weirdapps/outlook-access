// src/commands/move-mail.ts
//
// Move one or more messages to a destination folder. See:
//   - project-design.md §10.7 (CLI surface for move-mail)
//   - project-design.md §10.8 (Move semantics — returns NEW id)
//   - plan-002-folders.md §P5d
//   - refined-request-folders.md §5.4
//   - docs/research/outlook-v2-move-destination-alias.md
//
// Key invariants:
//
// 1. The `--to` spec is resolved exactly ONCE up front, before any /move
//    POST is issued (ADR-16: v2.0 `DestinationId` alias acceptance is
//    uncertain, so we always pre-resolve aliases / paths to a raw folder
//    id and pass that raw id as `DestinationId`).
//
// 2. Outlook's `POST /me/messages/{id}/move` returns a NEW message id
//    (project-design §10.8). The CLI surfaces the (sourceId, newId) pair
//    explicitly in `MoveMailResult.moved[]` so scripted users don't chain
//    stale ids downstream.
//
// 3. The loop is strictly single-threaded — v2.0 has no `$batch` and the
//    folder feature is NG10 (no concurrent moves).

import type { CliConfig } from '../config/config';
import { UpstreamError } from '../config/errors';
import { parseFolderSpec, resolveFolder } from '../folders/resolver';
import type { MoveEntry, MoveFailedEntry, MoveMailResult } from '../folders/types';
import type { OutlookClient } from '../http/outlook-client';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

// ---------------------------------------------------------------------------
// Public shapes
// ---------------------------------------------------------------------------

export interface MoveMailDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface MoveMailOptions {
  /** Destination folder. Well-known alias, display-name path, or `id:<raw>`. */
  to?: string;
  /** Tiebreaker for ambiguous path resolution. Default: false (exit 2 on ambiguity). */
  firstMatch?: boolean;
  /** If true, per-message failures are collected into `failed[]` instead of short-circuiting. */
  continueOnError?: boolean;
}

// Re-export for convenience — callers (cli.ts) import the result shape from
// `src/folders/types`, but the command's public contract is the `run()`
// signature below.
export type { MoveMailResult } from '../folders/types';

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/**
 * Move the given message ids into the destination folder identified by
 * `opts.to`.
 *
 * Control flow:
 *
 *   1. Validate argv (non-empty `messageIds`, `--to` provided).
 *   2. Load / refresh the session and construct the HTTP client.
 *   3. Resolve `--to` ONCE via `parseFolderSpec` + `resolveFolder` into a raw
 *      folder id (ADR-16). This single REST call also produces the
 *      `destination` block surfaced in the JSON output.
 *   4. Iterate source ids in order, issuing one `client.moveMessage` per id.
 *        - On success: push `{ sourceId, newId }` to `moved[]`.
 *        - On `ApiError` / `NetworkError`:
 *            * `--continue-on-error` → record in `failed[]` and continue.
 *            * otherwise             → short-circuit and throw (exit 5).
 *   5. Assemble and return the `MoveMailResult`. The top-level handler in
 *      `cli.ts` is responsible for mapping a non-empty `failed[]` to exit 5
 *      after emission.
 */
export async function run( // NOSONAR S3776 - batch move with error recovery
  deps: MoveMailDeps,
  messageIds: string[],
  opts: MoveMailOptions = {},
): Promise<MoveMailResult> {
  // ---- argv validation (raises exit 2) ----
  if (!Array.isArray(messageIds) || messageIds.length === 0) {
    throw new UsageError('move-mail: at least one <messageId> positional argument is required');
  }
  for (const id of messageIds) {
    if (typeof id !== 'string' || id.length === 0) {
      throw new UsageError('move-mail: <messageId> positional arguments must be non-empty strings');
    }
  }
  if (typeof opts.to !== 'string' || opts.to.length === 0) {
    throw new UsageError('move-mail: --to <spec> is required');
  }

  const firstMatch = opts.firstMatch === true;
  const continueOnError = opts.continueOnError === true;

  // ---- session + client ----
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  // ---- resolve destination ONCE (ADR-16) ----
  // parseFolderSpec turns the CLI string into a tagged spec; resolveFolder
  // produces a raw opaque id via the resolver (for 'id:...' it round-trips a
  // GET /me/MailFolders/{id}; for aliases it does the same; for paths it
  // walks). This is line where alias → raw-id normalization happens. The
  // resolver raises UsageError on ambiguity / UpstreamError on not-found.
  const spec = parseFolderSpec(opts.to);
  let destination;
  try {
    destination = await resolveFolder(client, spec, { firstMatch });
  } catch (err) {
    throw mapHttpError(err);
  }
  const destinationId = destination.Id;

  // ---- per-message loop ----
  const moved: MoveEntry[] = [];
  const failed: MoveFailedEntry[] = [];

  for (const sourceId of messageIds) {
    try {
      const result = await client.moveMessage(sourceId, destinationId);
      const newId =
        result !== null &&
        typeof result === 'object' &&
        typeof (result as { Id?: unknown }).Id === 'string'
          ? (result as { Id: string }).Id
          : '';
      if (newId.length === 0) {
        // Upstream returned 2xx with no Id. Treat as failure so the user is
        // not silently left without a new id to chain — R14 in plan-002 §7.
        const synthetic = new UpstreamError({
          code: 'UPSTREAM_HTTP_200',
          message: `move response for source '${sourceId}' is missing the new message Id.`,
        });
        if (continueOnError) {
          failed.push(toFailedEntry(sourceId, synthetic));
          continue;
        }
        throw synthetic;
      }
      moved.push({ sourceId, newId });
    } catch (err) {
      const mapped = mapHttpError(err);
      if (continueOnError) {
        failed.push(toFailedEntry(sourceId, mapped));
        continue;
      }
      // Short-circuit: first failure aborts the run. cli.ts maps this to
      // exit 5 (or 4 for AuthError).
      throw mapped;
    }
  }

  return {
    destination: {
      Id: destination.Id,
      Path: destination.Path,
      DisplayName: destination.DisplayName,
    },
    moved,
    failed,
    summary: {
      requested: messageIds.length,
      moved: moved.length,
      failed: failed.length,
    },
  };
}

// ---------------------------------------------------------------------------
// Helpers (inlined — plan-002 §P5d forbids a shared helper file across Wave-4)
// ---------------------------------------------------------------------------

/**
 * Translate a (already-mapped) error instance into a `MoveFailedEntry`. The
 * shape is stable per project-design §10.8:
 *   `{ sourceId, error: { code, httpStatus?, message? } }`
 *
 * We never include `cause` or any raw HTTP body fragment — `UpstreamError`
 * already runs its message through the redaction pipeline (see
 * `src/http/errors.ts` / `src/util/redact.ts`).
 */
function toFailedEntry(sourceId: string, err: unknown): MoveFailedEntry {
  // UpstreamError is the common case after `mapHttpError`.
  if (err instanceof UpstreamError) {
    return {
      sourceId,
      error: {
        code: err.code,
        httpStatus: err.httpStatus,
        message: err.message,
      },
    };
  }
  // Everything else — a best-effort generic shape. We avoid `any` and pull
  // only the fields we can reasonably type.
  const maybe = err as { code?: unknown; message?: unknown };
  const code =
    typeof maybe.code === 'string' && maybe.code.length > 0 ? maybe.code : 'UPSTREAM_UNKNOWN';
  const message = typeof maybe.message === 'string' ? maybe.message : String(err);
  return { sourceId, error: { code, message } };
}
