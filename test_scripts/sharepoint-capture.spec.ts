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

// Mimic a Playwright BrowserContext that captures requests at the context level.
// When `request` is provided, it is delivered to the registered 'request'
// listener as soon as the page navigates (mirroring real SharePoint emitting a
// Bearer call during initial load). When null, no request arrives, so the
// capture falls through to its timeout.
function makeFakeEnv(request: FakeRequest | null, cookies: FakeCookie[] = []) {
  let requestHandler: ((req: FakeRequest) => void) | undefined;
  const page = {
    goto: vi.fn().mockImplementation(async () => {
      if (request && requestHandler) requestHandler(request);
    }),
    close: vi.fn().mockResolvedValue(undefined),
  };
  const context = {
    newPage: vi.fn().mockResolvedValue(page),
    cookies: vi.fn().mockResolvedValue(cookies),
    on: vi.fn((event: string, handler: (req: FakeRequest) => void) => {
      if (event === 'request') requestHandler = handler;
    }),
    off: vi.fn(),
  };
  return { context, page };
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
    const { context, page } = makeFakeEnv(fakeRequest, cookies);

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

  it('captures a Bearer from an MCAS-proxied (*.mcas.ms) request', async () => {
    // MCAS Conditional Access App Control rewrites the SharePoint host to
    // "<fqdn>.mcas.ms"; the capture must still match it (regression test for
    // the headless/VPS timeout fix).
    const fakeRequest: FakeRequest = {
      url: () => 'https://nbg.sharepoint.com.mcas.ms/_api/web/lists',
      headers: () => ({ authorization: `Bearer ${FAKE_JWT}` }),
    };
    const { context } = makeFakeEnv(fakeRequest, []);

    const session = await captureSharepointFromContext(
      context as any,
      'nbg.sharepoint.com',
      30_000,
    );

    expect(session.bearer).toBe(FAKE_JWT);
    expect(session.host).toBe('nbg.sharepoint.com');
  });

  it('throws SHAREPOINT_INVALID_HOST for non-sharepoint hosts', async () => {
    const { context } = makeFakeEnv(null);
    await expect(
      captureSharepointFromContext(context as any, 'evil.example.com', 30_000),
    ).rejects.toThrow(SharepointCaptureError);
  });

  it('throws SHAREPOINT_TIMEOUT when no Bearer request arrives', async () => {
    // No request delivered → the capture falls through to its timeout. Use a
    // tiny timeout so the test stays fast.
    const { context } = makeFakeEnv(null);
    await expect(
      captureSharepointFromContext(context as any, 'tenant.sharepoint.com', 20),
    ).rejects.toThrow(/SharePoint Bearer/);
  });
});
