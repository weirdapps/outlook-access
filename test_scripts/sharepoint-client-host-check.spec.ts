// test_scripts/sharepoint-client-host-check.spec.ts
//
// The session cookies authorize the whole SharePoint tenant, so getBinary must
// never send them anywhere else: a mail's ReferenceAttachment SourceUrl is
// chosen by whoever sent the mail.
import { describe, it, expect, vi, afterEach } from 'vitest';
import { SharepointClient } from '../src/http/sharepoint-client';

const ORIGINAL_FETCH = global.fetch;

function okFetch() {
  return vi.fn(
    async () =>
      new Response(new Uint8Array([1]), {
        status: 200,
        headers: { 'content-type': 'application/octet-stream' },
      }),
  );
}

describe('SharepointClient refuses to send the session to a foreign host', () => {
  afterEach(() => {
    global.fetch = ORIGINAL_FETCH;
  });

  const tenant = () =>
    new SharepointClient({
      cookies: 'FedAuth=secret',
      timeoutMs: 30_000,
      host: 'contoso.sharepoint.com',
    });

  it.each([
    'https://evil.example/x.sharepoint.com/file.pdf',
    'https://contoso.sharepoint.com.evil.example/file.pdf',
    'https://evilsharepoint.com/file.pdf',
    'http://contoso.sharepoint.com/file.pdf',
    'https://x.sharepoint.com/file.pdf',
    'https://contoso.sharepoint.com@evil.example/file.pdf',
    'not a url',
  ])('refuses %s without calling fetch', async (url) => {
    const fetchMock = okFetch();
    global.fetch = fetchMock as unknown as typeof fetch;
    await expect(tenant().getBinary(url)).rejects.toThrow(/refusing|not a valid URL/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it.each([
    'https://contoso.sharepoint.com/sites/foo/Documents/Report.pdf',
    'https://contoso.sharepoint.com/:b:/s/foo/Eabc',
    'https://contoso-my.sharepoint.com/personal/someone/Documents/a.xlsx',
  ])('still fetches the own tenant: %s', async (url) => {
    const fetchMock = okFetch();
    global.fetch = fetchMock as unknown as typeof fetch;
    await tenant().getBinary(url);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('without a configured host, still refuses anything that is not https *.sharepoint.com', async () => {
    const fetchMock = okFetch();
    global.fetch = fetchMock as unknown as typeof fetch;
    const client = new SharepointClient({ cookies: 'FedAuth=secret', timeoutMs: 30_000 });
    await expect(client.getBinary('https://evil.example/file')).rejects.toThrow(/refusing/);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
