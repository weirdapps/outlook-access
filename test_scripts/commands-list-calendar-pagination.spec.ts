// test_scripts/commands-list-calendar-pagination.spec.ts
//
// list-calendar must return every event in the window, not the first page.
// It used a bare GET on /calendarview with no $top and never followed
// @odata.nextLink, and the server's default page is 10 events, so a week of
// meetings came back as its first 10.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import * as listCalendar from '../src/commands/list-calendar';
import type { CliConfig } from '../src/config/config';
import { createOutlookClient } from '../src/http/outlook-client';
import type { OutlookClient } from '../src/http/outlook-client';
import type { EventSummary } from '../src/http/types';
import type { SessionFile } from '../src/session/schema';

const JWT_SHAPED_TOKEN = 'aaaaaaaaaa.bbbbbbbbbb.cccccccccc';

function buildFakeSession(): SessionFile {
  return {
    version: 1,
    capturedAt: '2026-04-21T12:00:00.000Z',
    account: { upn: 'a@b.com', puid: '1', tenantId: 't' },
    bearer: {
      token: JWT_SHAPED_TOKEN,
      expiresAt: '2099-04-21T12:00:00.000Z',
      audience: 'https://outlook.office.com',
      scopes: ['Calendars.Read'],
    },
    cookies: [
      {
        name: 'C',
        value: 'v',
        domain: '.outlook.office.com',
        path: '/',
        expires: -1,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
    ],
    anchorMailbox: 'PUID:1@t',
  };
}

function buildFakeConfig(): CliConfig {
  return {
    httpTimeoutMs: 30_000,
    loginTimeoutMs: 300_000,
    chromeChannel: 'chrome',
    sessionFilePath: '/tmp/x',
    profileDir: '/tmp/p',
    tz: 'UTC',
    outputMode: 'json',
    listMailTop: 10,
    listMailFolder: 'Inbox',
    bodyMode: 'text',
    calFrom: 'now',
    calTo: 'now + 7d',
    quiet: true,
    noAutoReauth: false,
  } as CliConfig;
}

function makeEvent(id: string): EventSummary {
  return {
    Id: id,
    Subject: `Meeting ${id}`,
    Start: { DateTime: '2026-09-21T06:00:00.0000000', TimeZone: 'UTC' },
    End: { DateTime: '2026-09-21T07:00:00.0000000', TimeZone: 'UTC' },
  } as unknown as EventSummary;
}

function makeResponse(body: unknown): Response {
  const text = JSON.stringify(body);
  return {
    status: 200,
    ok: true,
    headers: new Headers({}),
    text: async () => text,
    json: async () => JSON.parse(text),
  } as unknown as Response;
}

describe('createOutlookClient.listCalendarView', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('follows @odata.nextLink until the window is exhausted', async () => {
    const page1 = Array.from({ length: 10 }, (_, i) => makeEvent(`e${i}`));
    const page2 = [makeEvent('e10'), makeEvent('e11')];
    const nextLink =
      'https://outlook.office.com/api/v2.0/me/calendarview?startDateTime=a&endDateTime=b&%24skip=10';
    fetchMock
      .mockResolvedValueOnce(makeResponse({ value: page1, '@odata.nextLink': nextLink }))
      .mockResolvedValueOnce(makeResponse({ value: page2 }));

    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const events = await client.listCalendarView({
      startDateTime: '2026-09-21T00:00:00.000Z',
      endDateTime: '2026-10-01T00:00:00.000Z',
    });

    expect(events.map((e) => e.Id)).toEqual([...page1, ...page2].map((e) => e.Id));
    expect(fetchMock).toHaveBeenCalledTimes(2);
    const [firstUrl] = fetchMock.mock.calls[0] as [string, unknown];
    expect(firstUrl).toContain('https://outlook.office.com/api/v2.0/me/calendarview');
    expect(firstUrl).toContain('startDateTime=');
    expect(firstUrl).toContain('%24top=');
    const [secondUrl] = fetchMock.mock.calls[1] as [string, unknown];
    expect(secondUrl).toBe(nextLink);
  });

  it('sends $top and keeps the window, $orderby and $select on the first page', async () => {
    fetchMock.mockResolvedValueOnce(makeResponse({ value: [makeEvent('e0')] }));
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await client.listCalendarView({
      startDateTime: '2026-09-21T00:00:00.000Z',
      endDateTime: '2026-10-01T00:00:00.000Z',
      $orderby: 'Start/DateTime asc',
      $select: 'Id,Subject,Start,End',
    });

    const [firstUrl] = fetchMock.mock.calls[0] as [string, unknown];
    const params = new URL(firstUrl).searchParams;
    expect(params.get('$top')).toBe('250');
    expect(params.get('startDateTime')).toBe('2026-09-21T00:00:00.000Z');
    expect(params.get('endDateTime')).toBe('2026-10-01T00:00:00.000Z');
    expect(params.get('$orderby')).toBe('Start/DateTime asc');
    expect(params.get('$select')).toBe('Id,Subject,Start,End');
  });

  it('delivers 1,000+ events even when the server serves 10 per page', async () => {
    const pages = 101;
    for (let p = 0; p < pages; p++) {
      const value = Array.from({ length: 10 }, (_, i) => makeEvent(`p${p}e${i}`));
      const body: Record<string, unknown> = { value };
      if (p < pages - 1) {
        body['@odata.nextLink'] =
          `https://outlook.office.com/api/v2.0/me/calendarview?%24skip=${(p + 1) * 10}`;
      }
      fetchMock.mockResolvedValueOnce(makeResponse(body));
    }
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    const events = await client.listCalendarView({ startDateTime: 'a', endDateTime: 'b' });

    expect(events).toHaveLength(1010);
    expect(fetchMock).toHaveBeenCalledTimes(pages);
  });

  it('refuses a nextLink that leaves outlook.office.com', async () => {
    fetchMock.mockResolvedValueOnce(
      makeResponse({
        value: [makeEvent('e0')],
        '@odata.nextLink': 'https://attacker.example.com/api/v2.0/me/calendarview?%24skip=1',
      }),
    );
    const client = createOutlookClient({
      session: buildFakeSession(),
      httpTimeoutMs: 5_000,
      noAutoReauth: false,
      onReauthNeeded: async () => buildFakeSession(),
    });

    await expect(
      client.listCalendarView({ startDateTime: 'a', endDateTime: 'b' }),
    ).rejects.toMatchObject({ code: 'UPSTREAM_PAGINATION_LIMIT' });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe('list-calendar run()', () => {
  it('returns every event the client pages through, with the window and $select', async () => {
    const events = Array.from({ length: 12 }, (_, i) => makeEvent(`e${i}`));
    const listCalendarView = vi.fn(async () => events);
    const client = { listCalendarView, get: vi.fn() } as unknown as OutlookClient;
    const session = buildFakeSession();

    const out = await listCalendar.run(
      {
        config: buildFakeConfig(),
        sessionPath: '/tmp/x',
        loadSession: async () => session,
        saveSession: async () => {},
        doAuthCapture: async () => session,
        createClient: () => client,
      },
      { from: '2026-09-21T00:00:00Z', to: '2026-10-01T00:00:00Z' },
    );

    expect(out).toHaveLength(12);
    const query = listCalendarView.mock.calls[0][0] as Record<string, string>;
    expect(query.startDateTime).toContain('2026-09-21');
    expect(query.endDateTime).toContain('2026-10-01');
    expect(query.$select).toContain('Start');
    expect((client as unknown as { get: ReturnType<typeof vi.fn> }).get).not.toHaveBeenCalled();
  });
});

describe('list-calendar run() error mapping through the paging client', () => {
  const fetchMock = vi.fn();

  beforeEach(() => {
    fetchMock.mockReset();
    vi.stubGlobal('fetch', fetchMock);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  function deps(noAutoReauth: boolean) {
    const session = buildFakeSession();
    return {
      config: buildFakeConfig(),
      sessionPath: '/tmp/x',
      loadSession: async () => session,
      saveSession: async () => {},
      doAuthCapture: async () => session,
      createClient: (s: SessionFile) =>
        createOutlookClient({
          session: s,
          httpTimeoutMs: 5_000,
          noAutoReauth,
          onReauthNeeded: async () => session,
        }),
    };
  }

  function errorResponse(status: number): Response {
    const text = JSON.stringify({ error: { code: 'X', message: 'boom' } });
    return {
      status,
      ok: false,
      headers: new Headers({}),
      text: async () => text,
      json: async () => JSON.parse(text),
    } as unknown as Response;
  }

  it('maps an HTTP 500 on a later page to UPSTREAM_HTTP_500', async () => {
    fetchMock
      .mockResolvedValueOnce(
        makeResponse({
          value: [makeEvent('e0')],
          '@odata.nextLink': 'https://outlook.office.com/api/v2.0/me/calendarview?%24skip=1',
        }),
      )
      .mockResolvedValueOnce(errorResponse(500));

    await expect(
      listCalendar.run(deps(false), { from: '2026-09-21T00:00:00Z', to: '2026-10-01T00:00:00Z' }),
    ).rejects.toMatchObject({ code: 'UPSTREAM_HTTP_500', httpStatus: 500 });
  });

  it('maps a 401 under noAutoReauth to AUTH_NO_REAUTH, as the single GET did', async () => {
    fetchMock.mockResolvedValueOnce(errorResponse(401));

    await expect(
      listCalendar.run(deps(true), { from: '2026-09-21T00:00:00Z', to: '2026-10-01T00:00:00Z' }),
    ).rejects.toMatchObject({ code: 'AUTH_NO_REAUTH' });
  });
});
