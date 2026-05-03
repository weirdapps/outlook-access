// test_scripts/commands-move-mail.spec.ts
//
// Command-level tests for `src/commands/move-mail.ts`.
//
// Scope:
//   - Alias pre-resolution: `--to <alias>` is resolved EXACTLY ONCE to a raw
//     folder id BEFORE any `moveMessage` call (ADR-16).
//   - Success path: (sourceId, newId) pairs in `moved[]`, `newId !== sourceId`.
//   - `--continue-on-error:false` short-circuits on first failure (subsequent
//     moveMessage calls are not issued).
//   - `--continue-on-error:true` continues; failure is collected in `failed[]`.
//   - Argv validation → UsageError (exit 2).
//   - `destination` block in the result carries {Id, Path, DisplayName} from
//     the one-shot resolve.
//
// No real HTTP. The OutlookClient is mocked with `Partial<OutlookClient>`.

import { describe, expect, it, vi } from 'vitest';

import { run as runMoveMail } from '../src/commands/move-mail';
import type { MoveMailDeps } from '../src/commands/move-mail';
import { UsageError } from '../src/commands/list-mail';
import type { CliConfig } from '../src/config/config';
import { ApiError } from '../src/http/errors';
import type { OutlookClient } from '../src/http/outlook-client';
import type { FolderSummary, MessageSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

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
      token: JWT_SHAPED_TOKEN,
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.Read', 'Mail.ReadWrite'],
    },
    cookies: [
      {
        name: 'SessionCookie',
        value: 'outlook-cookie-value',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
    ],
    anchorMailbox: 'PUID:1234567890@tenant-id-abc',
  };
}

function buildConfig(): CliConfig {
  return Object.freeze({
    httpTimeoutMs: 5_000,
    loginTimeoutMs: 60_000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/session.json',
    profileDir: '/tmp/profile',
    tz: 'UTC',
    outputMode: 'json',
    listMailTop: 10,
    listMailFolder: 'Inbox',
    bodyMode: 'text',
    calFrom: 'now',
    calTo: 'now + 7d',
    quiet: true,
    noAutoReauth: false,
  }) as CliConfig;
}

function buildDeps(client: Partial<OutlookClient>): MoveMailDeps {
  const session = buildFakeSession();
  return {
    config: buildConfig(),
    sessionPath: '/tmp/session.json',
    loadSession: async () => session,
    saveSession: async () => {
      /* no-op */
    },
    doAuthCapture: async () => session,
    createClient: () => client as OutlookClient,
  };
}

function inboxFolder(): FolderSummary {
  return {
    Id: 'inbox-raw-id',
    DisplayName: 'Inbox',
    WellKnownName: 'inbox',
  };
}

/**
 * Record every client call in a single ordered array so tests can assert the
 * relative ORDER of `getFolder` vs `moveMessage` (resolve-before-loop).
 */
interface CallLog {
  entries: Array<
    | { kind: 'getFolder'; arg: string }
    | { kind: 'moveMessage'; sourceId: string; destinationId: string }
    | { kind: 'listFolders'; parentId: string }
  >;
}

function buildMoveMessageResponse(newId: string): MessageSummary {
  return {
    Id: newId,
    Subject: 'moved',
    ReceivedDateTime: '2026-04-21T09:00:00Z',
    HasAttachments: false,
    IsRead: true,
    WebLink: 'https://outlook.office.com/owa/?ItemID=' + newId,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('move-mail command', () => {
  // -------------------------------------------------------------------------
  // Alias resolve-once-up-front + single-message happy path
  // -------------------------------------------------------------------------

  it('(1) single message, --to Inbox: alias resolved ONCE before any moveMessage; then moveMessage(msgId, rawId)', async () => {
    const log: CallLog = { entries: [] };

    const getFolder = vi.fn(async (idOrAlias: string) => {
      log.entries.push({ kind: 'getFolder', arg: idOrAlias });
      if (idOrAlias === 'Inbox') return inboxFolder();
      throw new Error(`unexpected getFolder(${idOrAlias})`);
    });

    const moveMessage = vi.fn(async (messageId: string, destinationFolderId: string) => {
      log.entries.push({
        kind: 'moveMessage',
        sourceId: messageId,
        destinationId: destinationFolderId,
      });
      return buildMoveMessageResponse(`${messageId}-NEW`);
    });

    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runMoveMail(deps, ['msg-1'], { to: 'Inbox' });

    // Exactly one resolve up-front.
    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(getFolder).toHaveBeenCalledWith('Inbox');

    // moveMessage was called exactly once, with the RAW id (not the alias).
    expect(moveMessage).toHaveBeenCalledTimes(1);
    expect(moveMessage).toHaveBeenCalledWith('msg-1', 'inbox-raw-id');

    // Strict ordering: getFolder(Inbox) MUST precede moveMessage(...).
    expect(log.entries[0]).toEqual({ kind: 'getFolder', arg: 'Inbox' });
    expect(log.entries[1]).toEqual({
      kind: 'moveMessage',
      sourceId: 'msg-1',
      destinationId: 'inbox-raw-id',
    });

    // Result shape sanity.
    expect(result.moved).toEqual([{ sourceId: 'msg-1', newId: 'msg-1-NEW' }]);
    expect(result.failed).toEqual([]);
    expect(result.summary).toEqual({ requested: 1, moved: 1, failed: 0 });
  });

  // -------------------------------------------------------------------------
  // 3 messages, all succeed
  // -------------------------------------------------------------------------

  it('(2) three messages all succeed: moved[] has 3 entries with newId != sourceId, failed[] empty', async () => {
    const log: CallLog = { entries: [] };

    const getFolder = vi.fn(async () => {
      log.entries.push({ kind: 'getFolder', arg: 'Inbox' });
      return inboxFolder();
    });

    const moveMessage = vi.fn(async (messageId: string, destinationFolderId: string) => {
      log.entries.push({
        kind: 'moveMessage',
        sourceId: messageId,
        destinationId: destinationFolderId,
      });
      return buildMoveMessageResponse(`${messageId}-NEW`);
    });

    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runMoveMail(deps, ['m1', 'm2', 'm3'], { to: 'Inbox' });

    // Resolve only once; three moves issued in order.
    expect(getFolder).toHaveBeenCalledTimes(1);
    expect(moveMessage).toHaveBeenCalledTimes(3);

    expect(result.moved.length).toBe(3);
    expect(result.failed.length).toBe(0);
    for (const entry of result.moved) {
      expect(entry.newId).not.toBe(entry.sourceId);
      expect(entry.newId.endsWith('-NEW')).toBe(true);
    }
    expect(result.moved).toEqual([
      { sourceId: 'm1', newId: 'm1-NEW' },
      { sourceId: 'm2', newId: 'm2-NEW' },
      { sourceId: 'm3', newId: 'm3-NEW' },
    ]);
    expect(result.summary).toEqual({ requested: 3, moved: 3, failed: 0 });

    // Order: resolve first, then move m1, m2, m3 in sequence.
    expect(log.entries.map((e) => e.kind)).toEqual([
      'getFolder',
      'moveMessage',
      'moveMessage',
      'moveMessage',
    ]);
  });

  // -------------------------------------------------------------------------
  // Continue-on-error semantics
  // -------------------------------------------------------------------------

  it('(3) first message fails + --continue-on-error:false throws immediately; subsequent moveMessage never called', async () => {
    const getFolder = vi.fn(async () => inboxFolder());
    const moveMessage = vi.fn(async (messageId: string) => {
      if (messageId === 'm1') {
        throw new ApiError({
          code: 'NOT_FOUND',
          message: 'message m1 is gone',
          httpStatus: 404,
          url: 'https://outlook.office.com/api/v2.0/me/messages/m1/move',
        });
      }
      return buildMoveMessageResponse(`${messageId}-NEW`);
    });

    const deps = buildDeps({ getFolder, moveMessage });

    await expect(
      runMoveMail(deps, ['m1', 'm2', 'm3'], {
        to: 'Inbox',
        continueOnError: false,
      }),
    ).rejects.toBeDefined();

    // Only m1 was attempted; the loop aborted before m2/m3.
    expect(moveMessage).toHaveBeenCalledTimes(1);
    expect(moveMessage).toHaveBeenCalledWith('m1', 'inbox-raw-id');
  });

  it('(4) first message fails + --continue-on-error:true: loop continues; failed.length=1, moved.length=2', async () => {
    const getFolder = vi.fn(async () => inboxFolder());
    const moveMessage = vi.fn(async (messageId: string) => {
      if (messageId === 'm1') {
        throw new ApiError({
          code: 'NOT_FOUND',
          message: 'message m1 is gone',
          httpStatus: 404,
          url: 'https://outlook.office.com/api/v2.0/me/messages/m1/move',
        });
      }
      return buildMoveMessageResponse(`${messageId}-NEW`);
    });

    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runMoveMail(deps, ['m1', 'm2', 'm3'], {
      to: 'Inbox',
      continueOnError: true,
    });

    // All three messages were attempted.
    expect(moveMessage).toHaveBeenCalledTimes(3);

    expect(result.moved.length).toBe(2);
    expect(result.failed.length).toBe(1);

    expect(result.moved).toEqual([
      { sourceId: 'm2', newId: 'm2-NEW' },
      { sourceId: 'm3', newId: 'm3-NEW' },
    ]);
    expect(result.failed[0].sourceId).toBe('m1');
    // The error is run through mapHttpError → UpstreamError; code surface is
    // a stable UPSTREAM_* string (see src/commands/list-mail.ts mapHttpError).
    expect(typeof result.failed[0].error.code).toBe('string');
    expect(result.failed[0].error.code.length).toBeGreaterThan(0);
    expect(result.failed[0].error.httpStatus).toBe(404);

    expect(result.summary).toEqual({ requested: 3, moved: 2, failed: 1 });
  });

  // -------------------------------------------------------------------------
  // Argv validation → UsageError (exit 2)
  // -------------------------------------------------------------------------

  it('(5) missing --to raises UsageError', async () => {
    const deps = buildDeps({});
    await expect(runMoveMail(deps, ['m1'], {})).rejects.toBeInstanceOf(UsageError);
  });

  it('(6) empty messageIds array raises UsageError', async () => {
    const deps = buildDeps({});
    await expect(runMoveMail(deps, [], { to: 'Inbox' })).rejects.toBeInstanceOf(UsageError);
  });

  // -------------------------------------------------------------------------
  // destination block shape
  // -------------------------------------------------------------------------

  it('(7) destination block in MoveMailResult has Id, Path, DisplayName from the resolve', async () => {
    const getFolder = vi.fn(async (idOrAlias: string) => {
      if (idOrAlias === 'Inbox') return inboxFolder();
      throw new Error(`unexpected getFolder(${idOrAlias})`);
    });
    const moveMessage = vi.fn(async (messageId: string) =>
      buildMoveMessageResponse(`${messageId}-NEW`),
    );

    const deps = buildDeps({ getFolder, moveMessage });

    const result = await runMoveMail(deps, ['m1'], { to: 'Inbox' });

    expect(result.destination.Id).toBe('inbox-raw-id');
    expect(result.destination.DisplayName).toBe('Inbox');
    // Resolver for a well-known alias materializes Path as the alias string.
    expect(result.destination.Path).toBe('Inbox');
  });
});
