// test_scripts/sharepoint-capture.spec.ts
//
// Tests captureSharepointFromContext with a mocked Playwright BrowserContext.
// Real Playwright is exercised in the manual smoke test (login --sharepoint-host).

import { describe, it, expect, vi } from 'vitest';
import {
  captureSharepointFromContext,
  SharepointCaptureError,
} from '../src/auth/sharepoint-capture';

// Build a fake-but-decodable JWT at runtime so the source has no hard-coded
// token literal (security scanners flag literal JWTs even in tests).
function b64url(s: string): string {
  return Buffer.from(s)
    .toString('base64')
    .replace(/=+$/, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
}
const FAR_FUTURE_EXP = 4102444800; // 2100-01-01T00:00:00Z
const FAKE_JWT_HEADER = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
const FAKE_JWT_PAYLOAD = b64url(
  JSON.stringify({
    exp: FAR_FUTURE_EXP,
    aud: 'https://nbg.sharepoint.com',
  }),
);
const FAKE_JWT_SIG = b64url('not-a-real-signature');
const FAKE_JWT = `${FAKE_JWT_HEADER}.${FAKE_JWT_PAYLOAD}.${FAKE_JWT_SIG}`;

interface FakeRequest {
  url: () => string;
  headers: () => Record<string, string>;
}

interface FakeCookie {
  name: string;
  value: string;
  domain: string;
}

function makeFakePage(request: FakeRequest) {
  return {
    goto: vi.fn().mockResolvedValue(undefined),
    waitForRequest: vi.fn().mockResolvedValue(request),
    close: vi.fn().mockResolvedValue(undefined),
  };
}

function makeFakeContext(page: ReturnType<typeof makeFakePage>, cookies: FakeCookie[] = []) {
  return {
    newPage: vi.fn().mockResolvedValue(page),
    cookies: vi.fn().mockResolvedValue(cookies),
  };
}

describe('captureSharepointFromContext', () => {
  it('extracts Bearer token + cookies for the SharePoint host', async () => {
    const fakeRequest: FakeRequest = {
      url: () => 'https://nbg.sharepoint.com/_api/web/lists',
      headers: () => ({ authorization: `Bearer ${FAKE_JWT}` }),
    };
    const cookies: FakeCookie[] = [
      { name: 'rtFa', value: 'abc', domain: '.sharepoint.com' },
      { name: 'FedAuth', value: 'def', domain: 'nbg.sharepoint.com' },
      { name: 'unrelated', value: 'xyz', domain: 'login.microsoftonline.com' },
    ];
    const page = makeFakePage(fakeRequest);
    const context = makeFakeContext(page, cookies);

    const session = await captureSharepointFromContext(
      context as any,
      'nbg.sharepoint.com',
      30_000,
    );

    expect(session.version).toBe(1);
    expect(session.host).toBe('nbg.sharepoint.com');
    expect(session.bearer).toBe(FAKE_JWT);
    expect(session.cookies).toContain('rtFa=abc');
    expect(session.cookies).toContain('FedAuth=def');
    expect(session.cookies).not.toContain('unrelated');
    expect(session.tokenExpiresAt).toBe(new Date(4102444800 * 1000).toISOString());
    expect(page.close).toHaveBeenCalled();
  });

  it('throws SHAREPOINT_INVALID_HOST for non-sharepoint hosts', async () => {
    const page = makeFakePage({ url: () => 'x', headers: () => ({}) });
    const context = makeFakeContext(page);
    await expect(
      captureSharepointFromContext(context as any, 'evil.example.com', 30_000),
    ).rejects.toThrow(SharepointCaptureError);
  });

  it('throws SHAREPOINT_TIMEOUT when waitForRequest rejects', async () => {
    const page = {
      goto: vi.fn().mockResolvedValue(undefined),
      waitForRequest: vi.fn().mockRejectedValue(new Error('Timeout 30000ms exceeded')),
      close: vi.fn().mockResolvedValue(undefined),
    };
    const context = makeFakeContext(page as unknown as ReturnType<typeof makeFakePage>);
    await expect(
      captureSharepointFromContext(context as any, 'tenant.sharepoint.com', 30_000),
    ).rejects.toThrow(/SharePoint Bearer/);
  });
});
