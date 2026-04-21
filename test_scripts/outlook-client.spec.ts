// test_scripts/outlook-client.spec.ts
//
// Tests for src/http/outlook-client.ts — retry logic, header/cookie building,
// and error-mapping.
// Covers AC-401-RETRY and error-mapping per project-design.md §2.8.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createOutlookClient } from '../src/http/outlook-client';
import {
  ApiError,
  AuthError,
  NetworkError,
} from '../src/http/errors';
import type { Cookie, SessionFile } from '../src/session/schema';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * A JWT-shaped token used for validation. NOT scanned by `redactString`
 * because it's well under the 100-char threshold.
 */
const JWT_SHAPED_TOKEN =
  'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

/**
 * 100+ char base64url-looking token. `redactString` DOES scrub runs of this
 * length, so using it in `bearer.token` is only safe for unit tests — in the
 * wild the session module guards it.
 *
 * The test uses the shorter JWT-shaped token by default (see session fake
 * below) because session validation requires a 3-segment JWT shape. The long
 * token below is used only for the redaction test where we stub it into a
 * response body rather than a session.
 */
const LONG_REDACTABLE_TOKEN =
  'A'.repeat(50) + 'B'.repeat(50) + 'C'.repeat(30) + 'test-bearer-XXXX';

function buildFakeSession(
  overrides: Partial<SessionFile> = {},
): SessionFile {
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

/** Build a Response-like object suitable for our fetch stub. */
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

// ---------------------------------------------------------------------------
// Test suite
// ---------------------------------------------------------------------------

describe('createOutlookClient.get', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('(1) 200 path returns parsed JSON and builds correct headers', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );

    const session = buildFakeSession();
    const client = createOutlookClient({
      session,
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => session,
    });

    const result = await client.get<{ value: unknown[] }>(
      '/api/v2.0/me/messages',
    );

    expect(result).toEqual({ value: [] });
    expect(fetchMock).toHaveBeenCalledTimes(1);

    const [callUrl, callInit] = fetchMock.mock.calls[0] as [
      string,
      { method: string; headers: Record<string, string> },
    ];
    expect(callUrl).toBe('https://outlook.office.com/api/v2.0/me/messages');
    expect(callInit.method).toBe('GET');
    expect(callInit.headers.Authorization).toBe(`Bearer ${JWT_SHAPED_TOKEN}`);
    expect(callInit.headers['X-AnchorMailbox']).toBe(
      'PUID:1234567890@tenant-id-abc',
    );
    expect(callInit.headers.Accept).toBe('application/json');
    expect(callInit.headers.Cookie).toBe(
      'SessionCookie=outlook-cookie-value',
    );
  });

  it('(2) 401 with auto-reauth retries and returns success', async () => {
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({ status: 401, bodyText: 'unauthorized' }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: [{ id: 'm1' }] } }),
      );

    const originalSession = buildFakeSession();
    const newSession = buildFakeSession({
      bearer: {
        ...originalSession.bearer,
        token: 'new.new.new',
      },
    });

    const onReauthNeeded = vi.fn(async () => newSession);

    const client = createOutlookClient({
      session: originalSession,
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded,
    });

    const result = await client.get<{ value: unknown[] }>(
      '/api/v2.0/me/messages',
    );

    expect(result).toEqual({ value: [{ id: 'm1' }] });
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(onReauthNeeded).toHaveBeenCalledTimes(1);

    // Second fetch must use the NEW bearer token.
    const secondCall = fetchMock.mock.calls[1] as [
      string,
      { headers: Record<string, string> },
    ];
    expect(secondCall[1].headers.Authorization).toBe('Bearer new.new.new');
  });

  it('(3) 401 with noAutoReauth throws AuthError and does NOT call reauth', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 401, bodyText: 'unauthorized' }),
    );

    const onReauthNeeded = vi.fn(async () => buildFakeSession());
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: true,
      onReauthNeeded,
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof AuthError &&
          err.code === 'AUTH_NO_REAUTH' &&
          err.reason === 'NO_AUTO_REAUTH'
        );
      },
    );

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(onReauthNeeded).not.toHaveBeenCalled();
  });

  it('(4) 401 twice (retry fails) throws AuthError with AFTER_RETRY reason', async () => {
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({ status: 401, bodyText: 'unauthorized #1' }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 401, bodyText: 'unauthorized #2' }),
      );

    const onReauthNeeded = vi.fn(async () => buildFakeSession());
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded,
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof AuthError &&
          err.reason === 'AFTER_RETRY' &&
          err.code === 'AUTH_REJECTED'
        );
      },
    );

    expect(onReauthNeeded).toHaveBeenCalledTimes(1);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('(5) 404 → ApiError with code NOT_FOUND', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 404, bodyText: 'not found' }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.get('/api/v2.0/me/messages/xyz')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof ApiError &&
          err.code === 'NOT_FOUND' &&
          err.httpStatus === 404
        );
      },
    );
  });

  it('(6) 429 with Retry-After header → ApiError RATE_LIMITED', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 429,
        bodyText: 'too many requests',
        headers: { 'retry-after': '60' },
      }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof ApiError &&
          err.code === 'RATE_LIMITED' &&
          err.httpStatus === 429 &&
          err.message.includes('Retry-After: 60')
        );
      },
    );
  });

  it('(7) 500 → ApiError with code SERVER_ERROR', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 500, bodyText: 'kaboom' }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof ApiError &&
          err.code === 'SERVER_ERROR' &&
          err.httpStatus === 500
        );
      },
    );
  });

  it('(8) Network error (TypeError) → NetworkError with timedOut=false', async () => {
    fetchMock.mockRejectedValueOnce(new TypeError('fetch failed'));

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return (
          err instanceof NetworkError &&
          err.timedOut === false &&
          err.code === 'NETWORK'
        );
      },
    );
  });

  it('(9) AbortError → NetworkError with timedOut=true', async () => {
    const abortErr = new Error('operation timed out');
    abortErr.name = 'AbortError';
    fetchMock.mockRejectedValueOnce(abortErr);

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(client.get('/api/v2.0/me/messages')).rejects.toSatisfy(
      (err: unknown) => {
        return err instanceof NetworkError && err.timedOut === true;
      },
    );
  });

  it('(10) redaction: long base64-looking tokens echoed in error bodies are scrubbed', async () => {
    // NOTE: `redactString` scrubs >100 char base64-ish runs. A plain
    // human-readable string is NOT scrubbed. This test uses a long
    // base64-looking token (>100 chars) to exercise the production redaction
    // path. The session bearer is still a short JWT-shape (schema required),
    // so we can't assert the bearer itself is scrubbed inside error text.
    const leakedToken = LONG_REDACTABLE_TOKEN; // 146 chars, base64-ish
    expect(leakedToken.length).toBeGreaterThan(100);

    fetchMock.mockResolvedValueOnce(
      makeResponse({
        status: 500,
        bodyText: `Upstream echoed the bearer: ${leakedToken} end`,
      }),
    );

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    let caught: unknown = null;
    try {
      await client.get('/api/v2.0/me/messages');
    } catch (err) {
      caught = err;
    }
    expect(caught).toBeInstanceOf(ApiError);
    const apiErr = caught as ApiError;
    expect(apiErr.message.indexOf(leakedToken)).toBe(-1);
    expect(JSON.stringify({ m: apiErr.message, u: apiErr.url }).indexOf(leakedToken)).toBe(-1);
  });

  it('(11) cookie domain filtering: only outlook.office.com cookies are sent', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({ status: 200, body: { value: [] } }),
    );

    const cookies: Cookie[] = [
      {
        name: 'OutlookOne',
        value: 'ok-1',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'OutlookTwo',
        value: 'ok-2',
        domain: 'outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: false,
        secure: true,
        sameSite: 'Lax',
      },
      {
        name: 'LoginCookie',
        value: 'nope-login',
        domain: '.login.microsoftonline.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'Evil',
        value: 'nope-evil',
        domain: 'evil.example.com',
        path: '/',
        expires: -1,
        httpOnly: false,
        secure: false,
        sameSite: 'Lax',
      },
    ];
    const session = buildFakeSession({ cookies });

    const client = createOutlookClient({
      session,
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => session,
    });

    await client.get('/api/v2.0/me/messages');

    const call = fetchMock.mock.calls[0] as [
      string,
      { headers: Record<string, string> },
    ];
    const cookieHeader = call[1].headers.Cookie ?? '';
    expect(cookieHeader).toContain('OutlookOne=ok-1');
    expect(cookieHeader).toContain('OutlookTwo=ok-2');
    expect(cookieHeader).not.toContain('LoginCookie');
    expect(cookieHeader).not.toContain('nope-login');
    expect(cookieHeader).not.toContain('Evil');
    expect(cookieHeader).not.toContain('nope-evil');
  });

  it('(12) session mutation after reauth: subsequent call uses new session', async () => {
    // Sequence:
    //   call#1: 401       -> reauth swaps session
    //   call#1 retry: 200
    //   call#2: 200       -> MUST use the refreshed session (new token/cookie)
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({ status: 401, bodyText: 'unauth' }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: [] } }),
      )
      .mockResolvedValueOnce(
        makeResponse({ status: 200, body: { value: [] } }),
      );

    const originalSession = buildFakeSession({
      cookies: [
        {
          name: 'OldCookie',
          value: 'old-val',
          domain: '.outlook.office.com',
          path: '/',
          expires: -1,
          httpOnly: true,
          secure: true,
          sameSite: 'None',
        },
      ],
    });
    const freshSession = buildFakeSession({
      bearer: {
        token: 'fresh.fresh.fresh',
        expiresAt: '2099-04-21T12:00:00.000Z',
        audience: 'https://outlook.office.com',
        scopes: ['Mail.Read'],
      },
      cookies: [
        {
          name: 'NewCookie',
          value: 'new-val',
          domain: '.outlook.office.com',
          path: '/',
          expires: -1,
          httpOnly: true,
          secure: true,
          sameSite: 'None',
        },
      ],
    });

    const onReauthNeeded = vi.fn(async () => freshSession);

    const client = createOutlookClient({
      session: originalSession,
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded,
    });

    await client.get('/api/v2.0/me/messages');
    await client.get('/api/v2.0/me/MailFolders');

    expect(fetchMock).toHaveBeenCalledTimes(3);

    const thirdCall = fetchMock.mock.calls[2] as [
      string,
      { headers: Record<string, string> },
    ];
    expect(thirdCall[1].headers.Authorization).toBe('Bearer fresh.fresh.fresh');
    expect(thirdCall[1].headers.Cookie).toBe('NewCookie=new-val');
    expect(thirdCall[1].headers.Cookie).not.toContain('OldCookie');
  });
});
