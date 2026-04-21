// test_scripts/outlook-client-folders.spec.ts
//
// Tests for the folder-feature additions on `OutlookClient`
// (src/http/outlook-client.ts): `listFolders`, `getFolder`, `createFolder`,
// `moveMessage`, `listMessagesInFolder`, plus observable behavior of the
// private `listAll<T>` and `doPost` helpers (pagination, nextLink follow,
// caps, and 401-retry ride-through).
//
// Kept sibling to `outlook-client.spec.ts` to keep that file focused on the
// shared retry/header/cookie envelope.
//
// Covers project-design.md §10.4 (API additions) and §10.6 (error mapping).

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient } from '../src/http/outlook-client';
import { CollisionError } from '../src/http/errors';
import { UpstreamError } from '../src/config/errors';
import { MAX_FOLDERS_VISITED } from '../src/folders/types';
import type { SessionFile } from '../src/session/schema';
import type { FolderSummary, MessageSummary } from '../src/http/types';

// ---------------------------------------------------------------------------
// Fixtures (match outlook-client.spec.ts conventions)
// ---------------------------------------------------------------------------

const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(overrides: Partial<SessionFile> = {}): SessionFile {
  const base: SessionFile = {
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
      scopes: ['Mail.Read', 'Calendars.Read'],
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
  return { ...base, ...overrides };
}

function makeResponse(init: {
  status: number;
  body?: unknown;
  bodyText?: string;
  headers?: Record<string, string>;
}): Response {
  const status = init.status;
  const headersMap = new Headers(init.headers ?? {});
  const bodyText =
    init.bodyText !== undefined
      ? init.bodyText
      : init.body !== undefined
        ? JSON.stringify(init.body)
        : '';
  return {
    status,
    ok: status >= 200 && status < 300,
    headers: headersMap,
    text: async () => bodyText,
    json: async () => {
      if (!bodyText) return undefined;
      return JSON.parse(bodyText);
    },
  } as unknown as Response;
}

function makeFolder(id: string, displayName: string): FolderSummary {
  return {
    Id: id,
    DisplayName: displayName,
    ParentFolderId: 'parent-0',
    ChildFolderCount: 0,
    UnreadItemCount: 0,
    TotalItemCount: 0,
  };
}

function makeMessage(id: string, subject: string): MessageSummary {
  return {
    Id: id,
    Subject: subject,
    ReceivedDateTime: '2026-04-21T12:00:00Z',
    HasAttachments: false,
    IsRead: false,
    WebLink: `https://outlook.office.com/owa/?ItemID=${id}`,
  };
}

// ---------------------------------------------------------------------------
// listFolders
// ---------------------------------------------------------------------------

describe('createOutlookClient.listFolders', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) single-page response returns the folders verbatim', async () => {
    const f1 = makeFolder('f1', 'Alpha');
    const f2 = makeFolder('f2', 'Beta');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [f1, f2] } }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listFolders('parent-0');
    expect(result).toEqual([f1, f2]);
    expect(fetchMock).toHaveBeenCalledTimes(1);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toContain(
      'https://outlook.office.com/api/v2.0/me/MailFolders/parent-0/childfolders',
    );
    // listAll sets a default $top if the caller didn't override.
    expect(url).toContain('%24top=');
  });

  it('(2) paginated response concatenates all pages via @odata.nextLink', async () => {
    const p1 = [makeFolder('f1', 'Alpha'), makeFolder('f2', 'Beta')];
    const p2 = [makeFolder('f3', 'Gamma')];
    const nextLink =
      'https://outlook.office.com/api/v2.0/me/MailFolders/parent-0/childfolders?%24skip=2&%24top=250';

    fetchMock
      .mockResolvedValueOnce(
        makeResponse({
          status: 200,
          body: { value: p1, '@odata.nextLink': nextLink },
        }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: p2 } }),
      );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listFolders('parent-0');
    expect(result).toEqual([...p1, ...p2]);
    expect(fetchMock).toHaveBeenCalledTimes(2);

    // Second call must hit the nextLink verbatim.
    const secondCall = fetchMock.mock.calls[1] as [string, unknown];
    expect(secondCall[0]).toBe(nextLink);
  });

  it('(3) exceeds MAX_FOLDERS_VISITED → UpstreamError UPSTREAM_PAGINATION_LIMIT', async () => {
    // Produce one page with MAX_FOLDERS_VISITED + 1 items; the collector trips
    // the item cap before it would otherwise have added the last entry.
    const oversized: FolderSummary[] = [];
    for (let i = 0; i < MAX_FOLDERS_VISITED + 1; i++) {
      oversized.push(makeFolder(`f${i}`, `Folder ${i}`));
    }

    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: oversized } }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.listFolders('parent-0')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof UpstreamError &&
          err.code === 'UPSTREAM_PAGINATION_LIMIT'
        );
      },
    );
  });

  it('(4) 401 on first page triggers auto-reauth retry then succeeds', async () => {
    const f1 = makeFolder('f1', 'Alpha');
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({ status: 401, bodyText: 'unauthorized' }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: [f1] } }),
      );

    const newSession = buildFakeSession({
      bearer: {
        token: 'new.new.new',
        expiresAt: '2099-04-21T12:00:00.000Z',
        audience: 'https://outlook.office.com',
        scopes: ['Mail.Read'],
      },
    });
    const onReauthNeeded = vi.fn(async () => newSession);

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded,
    });

    const result = await client.listFolders('parent-0');
    expect(result).toEqual([f1]);
    expect(onReauthNeeded).toHaveBeenCalledTimes(1);
    expect(fetchMock).toHaveBeenCalledTimes(2);

    const secondCall = fetchMock.mock.calls[1] as [
      string,
      { headers: Record<string, string> },
    ];
    expect(secondCall[1].headers.Authorization).toBe('Bearer new.new.new');
  });
});

// ---------------------------------------------------------------------------
// getFolder
// ---------------------------------------------------------------------------

describe('createOutlookClient.getFolder', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) happy path with well-known alias "Inbox"', async () => {
    const inbox = makeFolder('inbox-id', 'Inbox');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: inbox }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.getFolder('Inbox');
    expect(result).toEqual(inbox);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toBe('https://outlook.office.com/api/v2.0/me/MailFolders/Inbox');
  });

  it('(2) happy path with raw opaque id', async () => {
    const rawId = 'AAMkAGI1234567890';
    const folder = makeFolder(rawId, 'Projects');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: folder }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.getFolder(rawId);
    expect(result).toEqual(folder);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toBe(
      `https://outlook.office.com/api/v2.0/me/MailFolders/${rawId}`,
    );
  });

  it('(3) 404 → UpstreamError UPSTREAM_FOLDER_NOT_FOUND', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 404, bodyText: 'not found' }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.getFolder('does-not-exist')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof UpstreamError &&
          err.code === 'UPSTREAM_FOLDER_NOT_FOUND' &&
          err.httpStatus === 404
        );
      },
    );
  });
});

// ---------------------------------------------------------------------------
// createFolder
// ---------------------------------------------------------------------------

describe('createOutlookClient.createFolder', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) parentId === "msgfolderroot" posts to top-level /MailFolders', async () => {
    const created = makeFolder('new-id', 'Projects');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 201, body: created }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.createFolder('msgfolderroot', 'Projects');
    expect(result).toEqual(created);

    const [url, init] = fetchMock.mock.calls[0] as [
      string,
      { method: string; body: string; headers: Record<string, string> },
    ];
    expect(url).toBe('https://outlook.office.com/api/v2.0/me/MailFolders');
    expect(init.method).toBe('POST');
    expect(init.headers['Content-Type']).toBe('application/json');
    expect(JSON.parse(init.body)).toEqual({ DisplayName: 'Projects' });
  });

  it('(2) non-root parentId posts to /MailFolders/{parentId}/childfolders', async () => {
    const created = makeFolder('child-id', 'Alpha');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 201, body: created }),
    );

    const parentId = 'AAMkparent123';
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.createFolder(parentId, 'Alpha');
    expect(result).toEqual(created);

    const [url, init] = fetchMock.mock.calls[0] as [
      string,
      { method: string; body: string },
    ];
    expect(url).toBe(
      `https://outlook.office.com/api/v2.0/me/MailFolders/${parentId}/childfolders`,
    );
    expect(init.method).toBe('POST');
    expect(JSON.parse(init.body)).toEqual({ DisplayName: 'Alpha' });
  });

  it('(3) happy path returns the FolderSummary from the response body', async () => {
    const created: FolderSummary = {
      Id: 'brand-new',
      DisplayName: 'Brand New',
      ParentFolderId: 'parent-0',
      ChildFolderCount: 0,
      UnreadItemCount: 0,
      TotalItemCount: 0,
    };
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 201, body: created }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.createFolder('parent-0', 'Brand New');
    expect(result).toEqual(created);
  });

  it('(4) 400 + ErrorFolderExists body → CollisionError FOLDER_ALREADY_EXISTS', async () => {
    const dupeBody = {
      error: {
        code: 'ErrorFolderExists',
        message:
          "A folder with the specified name already exists., Could not create folder 'Alpha'.",
      },
    };
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 400, body: dupeBody }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.createFolder('parent-0', 'Alpha')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof CollisionError &&
          err.code === 'FOLDER_ALREADY_EXISTS' &&
          err.parentId === 'parent-0' &&
          err.path === 'Alpha'
        );
      },
    );
  });

  it('(5) 409 + ErrorFolderExists body → CollisionError (some tenants)', async () => {
    const dupeBody = {
      error: {
        code: 'ErrorFolderExists',
        message:
          "A folder with the specified name already exists., Could not create folder 'Alpha'.",
      },
    };
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 409, body: dupeBody }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.createFolder('parent-0', 'Alpha')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof CollisionError &&
          err.code === 'FOLDER_ALREADY_EXISTS'
        );
      },
    );
  });

  it('(6) non-ErrorFolderExists 400 → UpstreamError (NOT a CollisionError)', async () => {
    const otherBody = {
      error: {
        code: 'ErrorInvalidDisplayName',
        message: 'The folder name contains invalid characters.',
      },
    };
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 400, body: otherBody }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    let caught: unknown = null;
    try {
      await client.createFolder('parent-0', 'bad/name');
    } catch (err) {
      caught = err;
    }
    expect(caught).not.toBeNull();
    expect(caught).not.toBeInstanceOf(CollisionError);
    expect(caught).toBeInstanceOf(UpstreamError);
    expect((caught as UpstreamError).code).toBe('UPSTREAM_HTTP_400');
  });
});

// ---------------------------------------------------------------------------
// moveMessage
// ---------------------------------------------------------------------------

describe('createOutlookClient.moveMessage', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) happy path: POST body carries DestinationId; response returned verbatim', async () => {
    const moved: MessageSummary = {
      Id: 'NEW-ID-AFTER-MOVE',
      Subject: 'Hello',
      ReceivedDateTime: '2026-04-21T12:00:00Z',
      HasAttachments: false,
      IsRead: false,
      WebLink: 'https://outlook.office.com/owa/?ItemID=NEW-ID-AFTER-MOVE',
    };
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 201, body: moved }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.moveMessage('OLD-MSG-ID', 'dest-folder-123');
    expect(result).toEqual(moved);
    // The moved message's Id must differ from the source id — we pass through
    // whatever the server returns.
    expect(result.Id).toBe('NEW-ID-AFTER-MOVE');
    expect(result.Id).not.toBe('OLD-MSG-ID');

    const [url, init] = fetchMock.mock.calls[0] as [
      string,
      { method: string; body: string; headers: Record<string, string> },
    ];
    expect(url).toBe(
      'https://outlook.office.com/api/v2.0/me/messages/OLD-MSG-ID/move',
    );
    expect(init.method).toBe('POST');
    expect(init.headers['Content-Type']).toBe('application/json');
    expect(JSON.parse(init.body)).toEqual({ DestinationId: 'dest-folder-123' });
  });

  it('(2) 404 on destination → UpstreamError', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 404,
        body: { error: { code: 'ErrorItemNotFound', message: 'not found' } },
      }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(
      client.moveMessage('msg-id', 'missing-folder'),
    ).rejects.toSatisfy((err: unknown) => {
      return (
        err instanceof UpstreamError &&
        err.httpStatus === 404 &&
        err.code === 'UPSTREAM_HTTP_404'
      );
    });
  });

  it('(3) DestinationId is passed through verbatim (caller pre-resolves aliases)', async () => {
    // The method MUST NOT rewrite 'Inbox' or any alias — the caller
    // (resolver) is responsible for pre-resolution. This test pins that
    // contract by sending a raw alias-looking string and asserting the
    // wire body still contains it unchanged.
    const moved = makeMessage('new-id', 'Subject');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 201, body: moved }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.moveMessage('some-msg', 'Inbox');

    const [, init] = fetchMock.mock.calls[0] as [
      string,
      { body: string },
    ];
    expect(JSON.parse(init.body)).toEqual({ DestinationId: 'Inbox' });
  });
});

// ---------------------------------------------------------------------------
// listMessagesInFolder
// ---------------------------------------------------------------------------

describe('createOutlookClient.listMessagesInFolder', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) happy path: URL encodes $top/$select/$orderby and returns .value', async () => {
    const m1 = makeMessage('msg-1', 'Subject 1');
    const m2 = makeMessage('msg-2', 'Subject 2');
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [m1, m2] } }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesInFolder('folder-123', {
      top: 10,
      select: ['Id', 'Subject', 'ReceivedDateTime'],
      orderBy: 'ReceivedDateTime desc',
    });
    expect(result).toEqual([m1, m2]);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    const parsed = new URL(url);
    expect(parsed.pathname).toBe(
      '/api/v2.0/me/MailFolders/folder-123/messages',
    );
    expect(parsed.searchParams.get('$top')).toBe('10');
    expect(parsed.searchParams.get('$select')).toBe(
      'Id,Subject,ReceivedDateTime',
    );
    expect(parsed.searchParams.get('$orderby')).toBe('ReceivedDateTime desc');
  });

  it('(2) returns only the first page when upstream response carries items', async () => {
    // `listMessagesInFolder` uses the plain GET path (not `listAll`), so a
    // response with `@odata.nextLink` is ignored and only page-1 items are
    // returned.
    const p1 = [makeMessage('m1', 'one'), makeMessage('m2', 'two')];
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: {
          value: p1,
          '@odata.nextLink':
            'https://outlook.office.com/api/v2.0/me/MailFolders/folder-123/messages?%24skip=2',
        },
      }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesInFolder('folder-123', {
      top: 2,
    });
    expect(result).toEqual(p1);
    // A single underlying GET — no pagination follow-through.
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('(3) no options still produces a well-formed request path', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesInFolder('folder-xyz', {});
    expect(result).toEqual([]);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    const parsed = new URL(url);
    expect(parsed.pathname).toBe(
      '/api/v2.0/me/MailFolders/folder-xyz/messages',
    );
    // With no options, no query-string keys should have been appended.
    expect(parsed.searchParams.get('$top')).toBeNull();
    expect(parsed.searchParams.get('$select')).toBeNull();
    expect(parsed.searchParams.get('$orderby')).toBeNull();
  });
});
