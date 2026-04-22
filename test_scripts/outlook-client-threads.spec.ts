// test_scripts/outlook-client-threads.spec.ts
//
// Tests for the thread + date-filter additions on `OutlookClient`:
//   - `listMessagesInFolder` now accepts `filter` (threaded into $filter)
//   - `listMessagesByConversation` (new method)
//
// Kept sibling to outlook-client-folders.spec.ts; same mocking style.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';
import type { MessageSummary } from '../src/http/types';

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
      scopes: ['Mail.Read'],
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

function makeResponse(init: { status: number; body: unknown }): Response {
  const headersMap = new Headers();
  const bodyText = JSON.stringify(init.body);
  return {
    status: init.status,
    ok: init.status >= 200 && init.status < 300,
    headers: headersMap,
    text: async () => bodyText,
    json: async () => JSON.parse(bodyText),
  } as unknown as Response;
}

function makeMessage(id: string, received: string): MessageSummary {
  return {
    Id: id,
    Subject: `Subject ${id}`,
    ReceivedDateTime: received,
    HasAttachments: false,
    IsRead: false,
    WebLink: `https://example.com/${id}`,
  };
}

describe('listMessagesInFolder — filter option', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('threads a raw $filter expression into the query string', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [makeMessage('m1', '2026-04-01T10:00:00Z')] } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesInFolder('Inbox', {
      top: 10,
      filter: "ReceivedDateTime ge 2026-04-01T00:00:00.000Z and ReceivedDateTime lt 2026-05-01T00:00:00.000Z",
    });

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toContain('%24filter=');
    // URL-encoded form of the literal single-quote-free filter
    expect(decodeURIComponent(url.replace(/\+/g, '%20'))).toContain(
      'ReceivedDateTime ge 2026-04-01T00:00:00.000Z and ReceivedDateTime lt 2026-05-01T00:00:00.000Z',
    );
  });

  it('omits $filter when option is not provided', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesInFolder('Inbox', { top: 5 });
    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).not.toContain('%24filter=');
  });
});

describe('listMessagesByConversation', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('builds the expected /messages URL with $filter=ConversationId eq and defaults $orderby to ReceivedDateTime asc', async () => {
    const msgs = [
      makeMessage('m1', '2026-03-01T09:00:00Z'),
      makeMessage('m2', '2026-03-01T10:00:00Z'),
    ];
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: msgs } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.listMessagesByConversation('CONV-ABC-123');
    expect(result).toEqual(msgs);
    expect(fetchMock).toHaveBeenCalledTimes(1);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    expect(url).toContain('https://outlook.office.com/api/v2.0/me/messages?');
    const decoded = decodeURIComponent(url.replace(/\+/g, '%20'));
    expect(decoded).toContain("ConversationId eq 'CONV-ABC-123'");
    expect(decoded).toContain('ReceivedDateTime asc');
  });

  it('honors custom orderBy and select', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesByConversation('CID', {
      orderBy: 'ReceivedDateTime desc',
      select: ['Id', 'Subject', 'Body'],
      top: 50,
    });

    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    expect(decoded).toContain('ReceivedDateTime desc');
    expect(decoded).toContain('Id,Subject,Body');
    expect(decoded).toContain('$top=50');
  });

  it("escapes single quotes inside the conversation id (OData ' → '')", async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listMessagesByConversation("weird'id");
    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    expect(decoded).toContain("ConversationId eq 'weird''id'");
  });

  it('throws when conversationId is empty', async () => {
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await expect(client.listMessagesByConversation('')).rejects.toThrow(
      /non-empty conversationId/,
    );
  });
});

describe('countMessagesInFolder', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('sends $count=true&$top=1&$select=Id and returns @odata.count as exact:true', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { '@odata.count': 4273, value: [makeMessage('m1', '2026-04-01T10:00:00Z')] },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.countMessagesInFolder('Inbox');
    expect(result.count).toBe(4273);
    expect(result.exact).toBe(true);

    const [url] = fetchMock.mock.calls[0] as [string, unknown];
    const decoded = decodeURIComponent(url.replace(/\+/g, '%20'));
    expect(decoded).toContain('$count=true');
    expect(decoded).toContain('$top=1');
    expect(decoded).toContain('$select=Id');
    expect(url).toContain('/MailFolders/Inbox/messages');
  });

  it('threads filter into the request', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { '@odata.count': 12, value: [] },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.countMessagesInFolder('AAMk-raw-id', {
      filter: 'ReceivedDateTime ge 2026-04-01T00:00:00Z',
    });
    const decoded = decodeURIComponent(
      (fetchMock.mock.calls[0] as [string, unknown])[0].toString().replace(/\+/g, '%20'),
    );
    expect(decoded).toContain('ReceivedDateTime ge 2026-04-01T00:00:00Z');
  });

  it('falls back to value.length with exact:false when server omits @odata.count', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { value: [makeMessage('m1', '2026-04-01T10:00:00Z')] }, // no @odata.count
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.countMessagesInFolder('Inbox');
    expect(result.count).toBe(1);
    expect(result.exact).toBe(false);
  });

  it('throws when folderId is empty', async () => {
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await expect(client.countMessagesInFolder('')).rejects.toThrow(
      /non-empty folderId/,
    );
  });
});
