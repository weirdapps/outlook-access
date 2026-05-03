// test_scripts/outlook-client-send.spec.ts
//
// Tests for the v1.3.0 send/draft additions on `OutlookClient`:
//   - sendMail (immediate via /me/sendmail)
//   - createDraft (POST /me/messages → {Id, WebLink, ConversationId})
//   - sendDraft (POST /me/messages/{id}/send)
//
// Mocking style mirrors outlook-client-threads.spec.ts.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient, type SendMailPayload } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';

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
      scopes: ['Mail.ReadWrite', 'Mail.Send'],
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

const SAMPLE_PAYLOAD: SendMailPayload = {
  Subject: 'unit test',
  Body: { ContentType: 'HTML', Content: '<p>hi</p>' },
  ToRecipients: [{ EmailAddress: { Address: 'bob@example.com' } }],
};

describe('sendMail (immediate)', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('POSTs to /me/sendmail with {Message, SaveToSentItems:true} by default', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 202 }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.sendMail(SAMPLE_PAYLOAD);

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('https://outlook.office.com/api/v2.0/me/sendmail');
    expect(init.method).toBe('POST');
    const body = JSON.parse(init.body as string);
    expect(body.SaveToSentItems).toBe(true);
    expect(body.Message.Subject).toBe('unit test');
    expect(body.Message.Body.ContentType).toBe('HTML');
    expect(body.Message.ToRecipients[0].EmailAddress.Address).toBe('bob@example.com');
  });

  it('honors saveToSentItems: false', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 202 }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.sendMail(SAMPLE_PAYLOAD, { saveToSentItems: false });
    const [, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(JSON.parse(init.body as string).SaveToSentItems).toBe(false);
  });

  it('throws when server returns 401 with --no-auto-reauth', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 401, body: { error: { message: 'unauthorized' } } }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: true,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.sendMail(SAMPLE_PAYLOAD)).rejects.toThrow();
  });

  it('throws on 400 with the redacted body content in error message', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 400,
        body: {
          error: {
            message: 'echo back: {"Body":{"ContentType":"HTML","Content":"<p>secret</p>"}}',
          },
        },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.sendMail(SAMPLE_PAYLOAD)).rejects.toThrow(/\[REDACTED-BODY\]/);
  });
});

describe('createDraft + sendDraft', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('createDraft POSTs to /me/messages and returns {Id, WebLink, ConversationId}', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: {
          Id: 'AAMk-draft-001',
          WebLink: 'https://outlook.office.com/mail/drafts/id/AAMk-draft-001',
          ConversationId: 'conv-001',
          Subject: 'unit test',
        },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.createDraft(SAMPLE_PAYLOAD);

    expect(result.Id).toBe('AAMk-draft-001');
    expect(result.WebLink).toContain('outlook.office.com');
    expect(result.ConversationId).toBe('conv-001');

    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('https://outlook.office.com/api/v2.0/me/messages');
    expect(url).not.toContain('/sendmail');
    expect(init.method).toBe('POST');
    const body = JSON.parse(init.body as string);
    // createDraft sends the payload directly (no SaveToSentItems wrapper)
    expect(body.Subject).toBe('unit test');
    expect(body.SaveToSentItems).toBeUndefined();
  });

  it('createDraft tolerates missing ConversationId in server response', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 201,
        body: { Id: 'AAMk-draft-002', WebLink: 'https://x' },
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const result = await client.createDraft(SAMPLE_PAYLOAD);
    expect(result.Id).toBe('AAMk-draft-002');
    expect(result.ConversationId).toBeUndefined();
  });

  it('sendDraft POSTs to /me/messages/{id}/send with empty body', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 202 }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.sendDraft('AAMk-draft-001');

    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('/me/messages/AAMk-draft-001/send');
    expect(init.method).toBe('POST');
    expect(init.body).toBe('{}');
  });

  it('sendDraft URL-encodes the message id', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ status: 202 }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.sendDraft('id+with/special=chars');
    const [url] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toContain('/me/messages/id%2Bwith%2Fspecial%3Dchars/send');
  });

  it('sendDraft throws synchronously on empty messageId', async () => {
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });
    await expect(client.sendDraft('')).rejects.toThrow(/non-empty messageId/);
  });
});
