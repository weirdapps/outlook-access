// test_scripts/redact.spec.ts
//
// Unit tests for src/util/redact.ts — redactHeaders, redactJwt, redactString.
// Covers AC-NO-SECRET-LEAK.

import { describe, expect, it } from 'vitest';

import { redactHeaders, redactJwt, redactString, redactMessageBodies } from '../src/util/redact';

describe('redactHeaders', () => {
  it('redacts Authorization case-insensitively', () => {
    const out = redactHeaders({ Authorization: 'Bearer tok.xyz.abc' });
    expect(out.Authorization).toBe('[REDACTED]');
    const out2 = redactHeaders({ authorization: 'Bearer tok.xyz.abc' });
    expect(out2.authorization).toBe('[REDACTED]');
    const out3 = redactHeaders({ AUTHORIZATION: 'Bearer tok.xyz.abc' });
    expect(out3.AUTHORIZATION).toBe('[REDACTED]');
  });

  it('redacts Cookie and Set-Cookie headers', () => {
    const out = redactHeaders({
      Cookie: 'SIDEKICK=abc123; path=/',
      'Set-Cookie': 'FAKE=1; HttpOnly',
    });
    expect(out.Cookie).toBe('[REDACTED]');
    expect(out['Set-Cookie']).toBe('[REDACTED]');
  });

  it('redacts X-MS-*-Token headers (wildcard)', () => {
    const out = redactHeaders({
      'X-MS-MicrosoftGraph-Token': 'tok1',
      'x-ms-auth-token': 'tok2',
      'X-MS-Other-Token': 'tok3',
    });
    expect(out['X-MS-MicrosoftGraph-Token']).toBe('[REDACTED]');
    expect(out['x-ms-auth-token']).toBe('[REDACTED]');
    expect(out['X-MS-Other-Token']).toBe('[REDACTED]');
  });

  it('does not redact X-AnchorMailbox or other non-sensitive headers', () => {
    const out = redactHeaders({
      'X-AnchorMailbox': 'PUID:1234@tid',
      Accept: 'application/json',
      'User-Agent': 'outlook-cli',
    });
    expect(out['X-AnchorMailbox']).toBe('PUID:1234@tid');
    expect(out.Accept).toBe('application/json');
    expect(out['User-Agent']).toBe('outlook-cli');
  });

  it('returns a shallow copy (does not mutate the input object)', () => {
    const input = { Authorization: 'secret' };
    const out = redactHeaders(input);
    expect(input.Authorization).toBe('secret');
    expect(out.Authorization).toBe('[REDACTED]');
    expect(out).not.toBe(input);
  });
});

describe('redactJwt', () => {
  it('returns <first10>...<last5> for a long-enough string', () => {
    const s = '0123456789ABCDEFGHIJKLMNO';
    const out = redactJwt(s);
    expect(out).toBe('0123456789...KLMNO');
    // Sanity: neither the middle chars nor the full string appear in the output.
    expect(out).not.toContain('ABCDEF');
  });

  it('returns [REDACTED] placeholder for short tokens (< 16 chars)', () => {
    expect(redactJwt('short')).toBe('[REDACTED]');
    expect(redactJwt('')).toBe('[REDACTED]');
    expect(redactJwt('0123456789abcde')).toBe('[REDACTED]'); // length 15
  });
});

describe('redactString', () => {
  it('scrubs long base64/base64url runs (>100 chars)', () => {
    const longToken = 'A'.repeat(150);
    const input = `Error: failed to process ${longToken} at server`;
    const out = redactString(input);
    expect(out).toContain('[REDACTED]');
    expect(out).not.toContain(longToken);
    expect(out).toContain('Error: failed to process');
    expect(out).toContain('at server');
  });

  it('leaves short/regular strings alone', () => {
    expect(redactString('Hello, world!')).toBe('Hello, world!');
    expect(redactString('Request failed with status 403')).toBe('Request failed with status 403');
  });

  it('handles empty input', () => {
    expect(redactString('')).toBe('');
  });

  it('scrubs a JWT-shaped triple of base64url when length exceeds threshold', () => {
    const longJwt = 'A'.repeat(40) + '.' + 'B'.repeat(40) + '.' + 'C'.repeat(40);
    const out = redactString(longJwt);
    expect(out).toBe('[REDACTED]');
  });

  it('redacts message body content from echoed-back send-mail JSON', () => {
    const echo =
      '{"Subject":"test","Body":{"ContentType":"HTML","Content":"<p>secret content</p>"}}';
    const out = redactString(echo);
    expect(out).toContain('"Subject":"test"');
    expect(out).toContain('"ContentType":"HTML"');
    expect(out).not.toContain('secret content');
    expect(out).toContain('"Content":"[REDACTED-BODY]"');
  });

  it('redacts HtmlBody/TextBody flat field shapes', () => {
    const out1 = redactString('{"HtmlBody":"<b>secret</b>"}');
    expect(out1).toContain('"HtmlBody":"[REDACTED-BODY]"');
    expect(out1).not.toContain('secret');

    const out2 = redactString('{"TextBody":"the plain text body"}');
    expect(out2).toContain('"TextBody":"[REDACTED-BODY]"');
    expect(out2).not.toContain('plain text');
  });
});

describe('redactMessageBodies', () => {
  it('handles empty input', () => {
    expect(redactMessageBodies('')).toBe('');
  });

  it('leaves non-body content untouched', () => {
    expect(redactMessageBodies('{"Subject":"hi","From":"a@x.com"}')).toBe(
      '{"Subject":"hi","From":"a@x.com"}',
    );
  });

  it('preserves Subject and ContentType when redacting Body.Content', () => {
    const input =
      '{"Message":{"Subject":"S","Body":{"ContentType":"Text","Content":"body content here"}}}';
    const out = redactMessageBodies(input);
    expect(out).toContain('"Subject":"S"');
    expect(out).toContain('"ContentType":"Text"');
    expect(out).toContain('"Content":"[REDACTED-BODY]"');
    expect(out).not.toContain('body content here');
  });
});
