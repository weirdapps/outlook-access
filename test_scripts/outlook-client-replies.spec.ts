// test_scripts/outlook-client-replies.spec.ts
//
// Tests for the v1.4.0 reply/forward additions on `OutlookClient`:
//   - getMessage (single-message GET with $select)
//   - updateMessage (PATCH /me/messages/{id})
//   - createReply / createReplyAll / createForward (POST /me/messages/{id}/createX)

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';

const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: { upn: 'me@nbg.gr', puid: 'p', tenantId: 't' },
    bearer: {
      token: JWT_SHAPED_TOKEN,
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Mail.ReadWrite', 'Mail.Send'],
    },
    cookies: [],
    anchorMailbox: 'PUID:p@t',
  };
}

function makeResponse(init: { status: number; body?: unknown }): Response {
  const headersMap = new Headers();
  const bodyText = init.body === undefined ? '' : JSON.stringify(init.body);
  return {
    status: init.status,
    ok: init.status >= 200 && init.status < 300,
    headers: headersMap,
    text: async () => bodyText,
    json: async () => (bodyText ? JSON.parse(bodyText) : undefined),
  } as unknown as Response;
}

function newClient() {
  return createOutlookClient({
    session: buildFakeSession(),
    httpTimeoutMs: 5000,
    noAutoReauth: false,
    onReauthNeeded: async () => buildFakeSession(),
  });
}

describe('getMessage', () => {
  const fetchMock = vi.fn();
  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('GETs /me/messages/{id} with default $select projection', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { Id: 'AAMk-1', Subject: 'hi', WebLink: 'https://x' },
      }),
    );
    const result = await newClient().getMessage('AAMk-1');
    expect(result.Id).toBe('AAMk-1');
    expect(result.Subject).toBe('hi');
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(init.method).toBe('GET');
    expect(url).toContain('/me/messages/AAMk-1?');
    const decoded = decodeURIComponent(url.replace(/\+/g, '%20'));
    expect(decoded).toContain('$select=');
    expect(decoded).toContain('Body');
    expect(decoded).toContain('From');
  });

  it('honors custom select projection', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { Id: 'AAMk-2', Subject: 'h' } }),
    );
    await newClient().getMessage('AAMk-2', { select: ['Id', 'Subject'] });
    const [url] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(decodeURIComponent(url.replace(/\+/g, '%20'))).toContain(
      '$select=Id,Subject',
    );
  });

  it('rejects empty messageId synchronously', async () => {
    await expect(newClient().getMessage('')).rejects.toThrow(/non-empty messageId/);
  });
});

describe('updateMessage (PATCH)', () => {
  const fetchMock = vi.fn();
  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('PATCHes /me/messages/{id} with the supplied patch body', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 200,
        body: { Id: 'AAMk-3', Subject: 'updated' },
      }),
    );
    const patch = {
      Subject: 'updated',
      Body: { ContentType: 'HTML' as const, Content: '<p>new body</p>' },
    };
    const result = await newClient().updateMessage('AAMk-3', patch);
    expect(result.Subject).toBe('updated');
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(init.method).toBe('PATCH');
    expect(url).toContain('/me/messages/AAMk-3');
    expect(url).not.toContain('?');
    const body = JSON.parse(init.body as string);
    expect(body.Subject).toBe('updated');
    expect(body.Body.Content).toBe('<p>new body</p>');
  });

  it('sets Content-Type: application/json on PATCH (body-bearing method)', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 200, body: { Id: 'x' } }));
    await newClient().updateMessage('x', { Subject: 's' });
    const [, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    const headers = init.headers as Record<string, string>;
    expect(headers['Content-Type']).toBe('application/json');
  });

  it('rejects empty messageId synchronously', async () => {
    await expect(newClient().updateMessage('', {})).rejects.toThrow(
      /non-empty messageId/,
    );
  });
});

describe('createReply / createReplyAll / createForward', () => {
  const fetchMock = vi.fn();
  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('createReply POSTs to /me/messages/{id}/createReply with empty body', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: {
          Id: 'AAMk-reply-1',
          WebLink: 'https://x',
          Subject: 'RE: original',
          Body: { ContentType: 'HTML', Content: '<div>auto-quoted original</div>' },
          ToRecipients: [{ EmailAddress: { Address: 'sender@x.com' } }],
        },
      }),
    );
    const result = await newClient().createReply('AAMk-source-1');
    expect(result.Id).toBe('AAMk-reply-1');
    expect(result.Subject).toBe('RE: original');
    expect(result.Body.Content).toContain('auto-quoted');
    expect(result.ToRecipients).toHaveLength(1);

    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(init.method).toBe('POST');
    expect(url).toContain('/me/messages/AAMk-source-1/createReply');
    expect(init.body).toBe('{}');
  });

  it('createReplyAll POSTs to /createReplyAll endpoint', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: {
          Id: 'AAMk-replyall-1',
          WebLink: 'https://x',
          Subject: 'RE: original',
          Body: { ContentType: 'HTML', Content: '<div>quoted</div>' },
          ToRecipients: [
            { EmailAddress: { Address: 'a@x.com' } },
            { EmailAddress: { Address: 'b@y.com' } },
          ],
          CcRecipients: [{ EmailAddress: { Address: 'c@z.com' } }],
        },
      }),
    );
    const result = await newClient().createReplyAll('AAMk-source-2');
    expect(result.ToRecipients).toHaveLength(2);
    expect(result.CcRecipients).toHaveLength(1);
    const [url] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('/me/messages/AAMk-source-2/createReplyAll');
  });

  it('createForward POSTs to /createForward; recipients pre-empty', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: {
          Id: 'AAMk-fwd-1',
          WebLink: 'https://x',
          Subject: 'FW: original',
          Body: { ContentType: 'HTML', Content: '<div>quoted</div>' },
          ToRecipients: [],
        },
      }),
    );
    const result = await newClient().createForward('AAMk-source-3');
    expect(result.Subject).toBe('FW: original');
    expect(result.ToRecipients).toEqual([]);
    const [url] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('/me/messages/AAMk-source-3/createForward');
  });

  it('URL-encodes special chars in messageId', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: {
          Id: 'x',
          WebLink: 'https://x',
          Subject: 'RE: x',
          Body: { ContentType: 'HTML', Content: '' },
          ToRecipients: [],
        },
      }),
    );
    await newClient().createReply('id+with/special=chars');
    const [url] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('/me/messages/id%2Bwith%2Fspecial%3Dchars/createReply');
  });

  it('all three reject empty messageId synchronously', async () => {
    const c = newClient();
    await expect(c.createReply('')).rejects.toThrow(/createReply requires a non-empty messageId/);
    await expect(c.createReplyAll('')).rejects.toThrow(/createReplyAll requires a non-empty messageId/);
    await expect(c.createForward('')).rejects.toThrow(/createForward requires a non-empty messageId/);
  });
});
