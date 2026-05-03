// test_scripts/sharepoint-client.spec.ts
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { SharepointClient, SharepointHttpError } from '../src/http/sharepoint-client';

const ORIGINAL_FETCH = global.fetch;

describe('SharepointClient', () => {
  beforeEach(() => {
    // each test re-stubs fetch
  });
  afterEach(() => {
    global.fetch = ORIGINAL_FETCH;
  });

  it('attaches Bearer + Cookie headers to GET', async () => {
    const fetchMock = vi.fn(
      async () =>
        new Response(new Uint8Array([1, 2, 3, 4, 5, 6]), {
          status: 200,
          headers: {
            'content-type': 'application/octet-stream',
            'content-length': '6',
          },
        }),
    );
    global.fetch = fetchMock as unknown as typeof fetch;

    const client = new SharepointClient({
      bearer: 'redacted-bearer',
      cookies: 'rtFa=a; FedAuth=b',
      timeoutMs: 30_000,
    });
    const result = await client.getBinary('https://nbg.sharepoint.com/path/file.pdf');

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe('https://nbg.sharepoint.com/path/file.pdf');
    expect((init as RequestInit).headers).toMatchObject({
      Authorization: 'Bearer redacted-bearer',
      Cookie: 'rtFa=a; FedAuth=b',
    });
    expect(result.bytes.length).toBe(6);
    expect(result.size).toBe(6);
  });

  it('throws SharepointHttpError on 404 with status preserved', async () => {
    global.fetch = vi.fn(async () => new Response('', { status: 404 })) as unknown as typeof fetch;
    const client = new SharepointClient({ bearer: 'x', cookies: '', timeoutMs: 30_000 });
    let err: unknown;
    try {
      await client.getBinary('https://x.sharepoint.com/missing');
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(SharepointHttpError);
    expect((err as SharepointHttpError).status).toBe(404);
  });

  it('throws SharepointHttpError on 401', async () => {
    global.fetch = vi.fn(async () => new Response('', { status: 401 })) as unknown as typeof fetch;
    const client = new SharepointClient({ bearer: 'x', cookies: '', timeoutMs: 30_000 });
    await expect(client.getBinary('https://x.sharepoint.com/file')).rejects.toThrow(
      SharepointHttpError,
    );
  });

  it('parses filename from Content-Disposition (quoted)', async () => {
    global.fetch = vi.fn(
      async () =>
        new Response('hi', {
          status: 200,
          headers: { 'content-disposition': 'attachment; filename="report.pdf"' },
        }),
    ) as unknown as typeof fetch;
    const client = new SharepointClient({ bearer: 'x', cookies: '', timeoutMs: 30_000 });
    const result = await client.getBinary('https://x.sharepoint.com/r');
    expect(result.filename).toBe('report.pdf');
  });

  it('parses filename from Content-Disposition (RFC 5987 UTF-8)', async () => {
    global.fetch = vi.fn(
      async () =>
        new Response('hi', {
          status: 200,
          headers: {
            'content-disposition': "attachment; filename*=UTF-8''Q1%20Report.pdf",
          },
        }),
    ) as unknown as typeof fetch;
    const client = new SharepointClient({ bearer: 'x', cookies: '', timeoutMs: 30_000 });
    const result = await client.getBinary('https://x.sharepoint.com/r');
    expect(result.filename).toBe('Q1 Report.pdf');
  });

  it('omits Cookie header when cookies string is empty', async () => {
    const fetchMock = vi.fn(async () => new Response('x', { status: 200 }));
    global.fetch = fetchMock as unknown as typeof fetch;
    const client = new SharepointClient({ bearer: 'x', cookies: '', timeoutMs: 30_000 });
    await client.getBinary('https://x.sharepoint.com/y');
    const [, init] = fetchMock.mock.calls[0];
    expect((init as RequestInit).headers).not.toHaveProperty('Cookie');
  });
});
