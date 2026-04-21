// test_scripts/jwt.spec.ts
//
// Unit tests for src/auth/jwt.ts — decodeJwt.

import { describe, expect, it } from 'vitest';

import { decodeJwt } from '../src/auth/jwt';

/** Encode a string into base64url (URL-safe, no padding). */
function toBase64Url(input: string): string {
  return Buffer.from(input, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

/** Build a synthetic JWT from a claims object. Signature is arbitrary bytes. */
function buildJwt(payload: Record<string, unknown>): string {
  const header = toBase64Url(JSON.stringify({ alg: 'none', typ: 'JWT' }));
  const body = toBase64Url(JSON.stringify(payload));
  const sig = toBase64Url('test-signature-bytes');
  return `${header}.${body}.${sig}`;
}

describe('decodeJwt', () => {
  it('returns claims for a well-formed synthetic JWT', () => {
    const exp = Math.floor(Date.now() / 1000) + 3600;
    const token = buildJwt({
      exp,
      aud: 'https://outlook.office.com/',
      oid: 'abc-123',
      tid: 'tenant-1',
      upn: 'alice@contoso.com',
      scp: 'Mail.Read Mail.Send',
    });
    const claims = decodeJwt(token);
    expect(claims.exp).toBe(exp);
    expect(claims.aud).toBe('https://outlook.office.com/');
    expect(claims.oid).toBe('abc-123');
    expect(claims.tid).toBe('tenant-1');
    expect(claims.upn).toBe('alice@contoso.com');
    expect(claims.scp).toBe('Mail.Read Mail.Send');
  });

  it('strips the "Bearer " prefix before decoding', () => {
    const exp = Math.floor(Date.now() / 1000) + 60;
    const token = buildJwt({ exp, aud: 'aud-xyz' });
    const claims = decodeJwt('Bearer ' + token);
    expect(claims.aud).toBe('aud-xyz');
    expect(claims.exp).toBe(exp);
  });

  it('throws "invalid JWT" on only one segment', () => {
    expect(() => decodeJwt('onlyonesegment')).toThrowError(/invalid JWT/);
  });

  it('throws "invalid JWT" on only two segments', () => {
    expect(() => decodeJwt('header.payload')).toThrowError(/invalid JWT/);
  });

  it('throws "invalid JWT" on empty string', () => {
    expect(() => decodeJwt('')).toThrowError(/invalid JWT/);
  });

  it('handles URL-safe base64 chars (- and _) plus missing padding', () => {
    // Force a payload whose standard base64 would contain "+" and "/"
    // characters; base64url replaces them with "-" and "_" and strips "=" pad.
    const exp = 1234567890;
    // Use a payload string carefully designed to include "+" and "/" when
    // encoded as standard base64. The characters ">?" in JSON (hex 3E3F) give
    // the bits that produce "+/" in standard base64. We include them inside a
    // string value in an optional claim.
    const payload = { exp, aud: 'a', marker: '>?>?>?>?' };

    const raw = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64');
    expect(raw).toMatch(/[+/]/);
    const urlSafe = raw.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    expect(urlSafe).not.toMatch(/=/);

    const header = toBase64Url(JSON.stringify({ alg: 'none' }));
    const token = `${header}.${urlSafe}.${toBase64Url('sig')}`;

    const claims = decodeJwt(token);
    expect(claims.exp).toBe(exp);
    expect(claims.aud).toBe('a');
    expect(claims['marker']).toBe('>?>?>?>?');
  });

  it('throws on three segments where payload is not valid base64/json', () => {
    // Payload "%%%%%" decodes to garbage that is not valid JSON.
    const token = 'aaa.%%%%%%.bbb';
    expect(() => decodeJwt(token)).toThrowError(/invalid JWT/);
  });

  it('throws when exp claim is missing', () => {
    const header = toBase64Url(JSON.stringify({ alg: 'none' }));
    const body = toBase64Url(JSON.stringify({ aud: 'x' }));
    const token = `${header}.${body}.${toBase64Url('sig')}`;
    expect(() => decodeJwt(token)).toThrowError(/invalid JWT/);
  });
});
