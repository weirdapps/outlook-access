// test_scripts/sharepoint-schema.spec.ts
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { describe, it, expect } from 'vitest';
import {
  parseSharepointSession,
  serializeSharepointSession,
  loadSharepointSession,
  saveSharepointSession,
  defaultSharepointSessionPath,
  SharepointSessionParseError,
  SharepointSession,
} from '../src/session/sharepoint-schema';

const SAMPLE: SharepointSession = {
  version: 1,
  host: 'nbg.sharepoint.com',
  bearer: 'eyJ.fake.token',
  cookies: 'rtFa=abc; FedAuth=def',
  capturedAt: '2026-04-22T14:00:00.000Z',
  tokenExpiresAt: '2026-04-22T22:00:00.000Z',
};

describe('sharepoint session schema', () => {
  it('round-trips a valid session', () => {
    const json = serializeSharepointSession(SAMPLE);
    expect(parseSharepointSession(json)).toEqual(SAMPLE);
  });

  it('accepts a cookie-only session (bearer omitted)', () => {
    // Cookie-auth tenants have no Bearer; the session is valid on cookies alone.
    const { bearer: _omit, ...cookieOnly } = SAMPLE;
    const json = JSON.stringify(cookieOnly);
    const parsed = parseSharepointSession(json);
    expect(parsed.bearer).toBeUndefined();
    expect(parsed.cookies).toBe(SAMPLE.cookies);
  });

  it('rejects a non-string bearer', () => {
    const broken = JSON.stringify({ ...SAMPLE, bearer: 123 });
    expect(() => parseSharepointSession(broken)).toThrow(SharepointSessionParseError);
    expect(() => parseSharepointSession(broken)).toThrow(/bearer/);
  });

  it('rejects unknown version', () => {
    const broken = JSON.stringify({ ...SAMPLE, version: 99 });
    expect(() => parseSharepointSession(broken)).toThrow(/version/);
  });

  it('rejects malformed JSON', () => {
    expect(() => parseSharepointSession('not json')).toThrow(/Invalid JSON/);
  });

  it('rejects null root', () => {
    expect(() => parseSharepointSession('null')).toThrow(/object/);
  });

  it('save then load round trip on disk with mode 0600', async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sp-schema-'));
    const file = path.join(tmp, 'sharepoint-session.json');
    await saveSharepointSession(file, SAMPLE);
    const loaded = await loadSharepointSession(file);
    expect(loaded).toEqual(SAMPLE);
    const stat = fs.statSync(file);
    // mode bits — file should be 0o600
    expect(stat.mode & 0o777).toBe(0o600);
  });

  it('loadSharepointSession returns null on missing file', async () => {
    const result = await loadSharepointSession('/tmp/definitely-does-not-exist-' + Date.now());
    expect(result).toBeNull();
  });

  it('defaultSharepointSessionPath returns ~/.outlook-cli/sharepoint-session.json', () => {
    expect(defaultSharepointSessionPath()).toBe(
      path.join(os.homedir(), '.outlook-cli', 'sharepoint-session.json'),
    );
  });
});
