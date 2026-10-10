// test_scripts/commands-delete-draft.spec.ts
//
// Command-level tests for `src/commands/delete-draft.ts`.
//
// Scope:
//   - Only messages the server reports as `IsDraft: true` are deleted.
//   - A non-draft is refused BEFORE any delete call, with no fallback.
//   - `--continue-on-error` collects failures in `failed[]`; without it the
//     first failure short-circuits.
//   - Argv validation → UsageError (exit 2).
//
// No real HTTP. The OutlookClient is mocked with `Partial<OutlookClient>`.

import { describe, expect, it, vi } from 'vitest';

import { run as runDeleteDraft } from '../src/commands/delete-draft';
import type { DeleteDraftDeps } from '../src/commands/delete-draft';
import { UsageError } from '../src/commands/list-mail';
import type { CliConfig } from '../src/config/config';
import { UpstreamError } from '../src/config/errors';
import type { GetMessageResult, OutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: { upn: 'alice@contoso.com', puid: '1234567890', tenantId: 'tenant-id-abc' },
    bearer: {
      token: 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc',
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.ReadWrite'],
    },
    cookies: [],
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

function buildDeps(client: Partial<OutlookClient>): DeleteDraftDeps {
  const session = buildFakeSession();
  return {
    config: buildConfig(),
    sessionPath: '/tmp/session.json',
    loadSession: async () => session,
    saveSession: async () => {},
    doAuthCapture: async () => session,
    createClient: () => client as OutlookClient,
  };
}

function msg(id: string, isDraft: boolean): GetMessageResult {
  return { Id: id, Subject: `subject ${id}`, IsDraft: isDraft };
}

describe('delete-draft', () => {
  it('deletes a message the server reports as a draft', async () => {
    const getMessage = vi.fn(async (id: string) => msg(id, true));
    const deleteMessage = vi.fn(async () => {});
    const res = await runDeleteDraft(buildDeps({ getMessage, deleteMessage }), ['d1']);

    expect(getMessage).toHaveBeenCalledWith('d1', { select: ['Id', 'Subject', 'IsDraft'] });
    expect(deleteMessage).toHaveBeenCalledWith('d1');
    expect(res.deleted).toEqual([{ id: 'd1', subject: 'subject d1' }]);
    expect(res.failed).toEqual([]);
    expect(res.summary).toEqual({ requested: 1, deleted: 1, failed: 0 });
  });

  it('refuses a sent or received message and never calls delete', async () => {
    const getMessage = vi.fn(async (id: string) => msg(id, false));
    const deleteMessage = vi.fn(async () => {});
    await expect(
      runDeleteDraft(buildDeps({ getMessage, deleteMessage }), ['m1']),
    ).rejects.toBeInstanceOf(UpstreamError);
    expect(deleteMessage).not.toHaveBeenCalled();
  });

  it('treats a missing IsDraft as not a draft', async () => {
    const getMessage = vi.fn(async (id: string) => ({ Id: id, Subject: 's' }));
    const deleteMessage = vi.fn(async () => {});
    const res = await runDeleteDraft(buildDeps({ getMessage, deleteMessage }), ['m1'], {
      continueOnError: true,
    });
    expect(deleteMessage).not.toHaveBeenCalled();
    expect(res.failed[0]).toMatchObject({ id: 'm1', error: { code: 'NOT_A_DRAFT' } });
  });

  it('with --continue-on-error deletes the drafts and reports the rest', async () => {
    const getMessage = vi.fn(async (id: string) => msg(id, id !== 'sent'));
    const deleteMessage = vi.fn(async () => {});
    const res = await runDeleteDraft(
      buildDeps({ getMessage, deleteMessage }),
      ['d1', 'sent', 'd2'],
      {
        continueOnError: true,
      },
    );
    expect(deleteMessage.mock.calls.map((c) => c[0])).toEqual(['d1', 'd2']);
    expect(res.failed.map((f) => f.id)).toEqual(['sent']);
    expect(res.summary).toEqual({ requested: 3, deleted: 2, failed: 1 });
  });

  it('without --continue-on-error stops at the first refusal', async () => {
    const getMessage = vi.fn(async (id: string) => msg(id, id !== 'sent'));
    const deleteMessage = vi.fn(async () => {});
    await expect(
      runDeleteDraft(buildDeps({ getMessage, deleteMessage }), ['sent', 'd1']),
    ).rejects.toBeInstanceOf(UpstreamError);
    expect(getMessage).toHaveBeenCalledTimes(1);
    expect(deleteMessage).not.toHaveBeenCalled();
  });

  it('collects a delete that fails upstream, keeping its code', async () => {
    const getMessage = vi.fn(async (id: string) => msg(id, true));
    const deleteMessage = vi.fn(async () => {
      throw new UpstreamError({ code: 'UPSTREAM_HTTP_404', message: 'gone', httpStatus: 404 });
    });
    const res = await runDeleteDraft(buildDeps({ getMessage, deleteMessage }), ['d1'], {
      continueOnError: true,
    });
    expect(res.failed[0]).toMatchObject({
      id: 'd1',
      error: { code: 'UPSTREAM_HTTP_404', httpStatus: 404 },
    });
  });

  it('maps a non-upstream error into failed[] with a generic code', async () => {
    const getMessage = vi.fn(async () => {
      throw new Error('socket hang up');
    });
    const res = await runDeleteDraft(buildDeps({ getMessage }), ['d1'], { continueOnError: true });
    expect(res.failed[0].id).toBe('d1');
    expect(res.failed[0].error.message).toMatch(/socket hang up/);
  });

  it('rejects an empty-string id', async () => {
    await expect(runDeleteDraft(buildDeps({}), [''])).rejects.toBeInstanceOf(UsageError);
  });

  it('rejects an empty id list', async () => {
    await expect(runDeleteDraft(buildDeps({}), [])).rejects.toBeInstanceOf(UsageError);
  });
});
