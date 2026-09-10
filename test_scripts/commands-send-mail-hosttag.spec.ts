// test_scripts/commands-send-mail-hosttag.spec.ts
//
// Unit tests for the per-host subject tag in src/commands/send-mail.ts.
//
// Three machines run this CLI against ONE mailbox and their reports arrived
// indistinguishable: on 2026-08-29 two nightly health mails landed a minute
// apart, contradicting each other, with nothing in either saying which box had
// produced it. The tag is scoped to SELF-ADDRESSED mail on purpose. Tagging
// unconditionally would stamp "[Pro]" on mail to colleagues, which is the one
// outcome worse than the confusion it fixes.
//
// Everything host-specific lives in ~/.outlook-cli/, never in this PUBLIC repo.

import { describe, it, expect, vi } from 'vitest';

import { run } from '../src/commands/send-mail';
import type { OutlookClient } from '../src/http/outlook-client';
import type { SessionFile } from '../src/session/schema';
import type { CliConfig } from '../src/config/config';

const MINIMAL_CONFIG = {
  httpTimeoutMs: 30000,
  loginTimeoutMs: 300000,
  chromeChannel: 'chrome',
  sessionFilePath: '/tmp/session.json',
  profileDir: '/tmp/profile',
  tz: 'UTC',
  outputMode: 'json',
  listMailTop: 10,
  listMailFolder: 'Inbox',
  bodyMode: 'text',
  calFrom: 'now',
  calTo: 'now + 7d',
  quiet: true,
  noAutoReauth: true,
} as unknown as CliConfig;

const SESSION: SessionFile = {
  version: 1,
  capturedAt: '2026-04-21T12:00:00.000Z',
  account: { upn: 'me@example.com', puid: 'p', tenantId: 't' },
  bearer: {
    token: 'x.y.z',
    expiresAt: '2099-04-21T12:00:00.000Z',
    audience: 'https://outlook.office.com',
    scopes: ['Mail.Send'],
  },
  cookies: [],
  anchorMailbox: 'PUID:p@t',
};

const HOME = '/tmp/fake-home';
const HOST_TAG = `${HOME}/.outlook-cli/host-tag`;
const SELF_ADDRS = `${HOME}/.outlook-cli/self-addresses`;
const BODY_PATH = '/tmp/b.html';
const BODY = { [BODY_PATH]: '<p>hi</p>' };

function makeDeps(fileMap: Record<string, string> = {}) {
  const client = {
    sendMail: vi.fn(async () => undefined),
    createDraft: vi.fn(async () => ({
      Id: 'AAMk-draft-001',
      WebLink: 'https://outlook.office.com/mail/drafts/id/AAMk-draft-001',
      ConversationId: 'conv-001',
    })),
  } as unknown as OutlookClient;

  const readFile = vi.fn(async (p: string) => {
    if (p in fileMap) return Buffer.from(fileMap[p] as string, 'utf-8');
    throw Object.assign(new Error(`ENOENT: ${p}`), { code: 'ENOENT' });
  });

  return {
    deps: {
      config: MINIMAL_CONFIG,
      sessionPath: '/tmp/session.json',
      loadSession: vi.fn(async () => SESSION),
      saveSession: vi.fn(async () => {}),
      doAuthCapture: vi.fn(async () => SESSION),
      createClient: vi.fn(() => client),
      activateOutlook: vi.fn(async () => undefined),
      readFile,
      homeDir: () => HOME,
    },
    client,
  };
}

const subjectOf = (client: OutlookClient) =>
  (client.createDraft as ReturnType<typeof vi.fn>).mock.calls[0]![0].Subject;

describe('send-mail host tag', () => {
  it('tags when every recipient is the authenticated user', async () => {
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: 'Pro\n' });
    await run(deps, { to: 'me@example.com', subject: 'nightly report', html: BODY_PATH });
    expect(subjectOf(client)).toBe('[Pro] nightly report');
  });

  it('does NOT tag mail addressed to anyone else', async () => {
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: 'Pro\n' });
    await run(deps, {
      to: 'colleague@work.example',
      subject: 'quarterly numbers',
      html: BODY_PATH,
    });
    expect(subjectOf(client)).toBe('quarterly numbers');
  });

  it('does NOT tag when self is only one of several recipients', async () => {
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: 'Pro\n' });
    await run(deps, {
      to: ['me@example.com', 'colleague@work.example'],
      subject: 'numbers',
      html: BODY_PATH,
    });
    expect(subjectOf(client)).toBe('numbers');
  });

  it('counts a configured alias as self', async () => {
    // The mail that prompted this goes to a FORWARDING ALIAS, not the UPN, so
    // a strict UPN comparison would miss the exact case it was written for.
    const { deps, client } = makeDeps({
      ...BODY,
      [HOST_TAG]: 'Neo\n',
      [SELF_ADDRS]: '# my own addresses\nalias@example.com\n\n',
    });
    await run(deps, { to: 'alias@example.com', subject: 'automation health', html: BODY_PATH });
    expect(subjectOf(client)).toBe('[Neo] automation health');
  });

  it('is idempotent: never double-tags an already-tagged subject', async () => {
    // The VPS wrapper at ~/scripts/outlook-cli prepends its own tag before it
    // calls this binary, so a subject legitimately arrives pre-tagged.
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: 'VPS\n' });
    await run(deps, { to: 'me@example.com', subject: '[VPS] already tagged', html: BODY_PATH });
    expect(subjectOf(client)).toBe('[VPS] already tagged');
  });

  it('does nothing at all when no host-tag file exists', async () => {
    // The default for every other user of this public repo: unchanged.
    const { deps, client } = makeDeps(BODY);
    await run(deps, { to: 'me@example.com', subject: 'untouched', html: BODY_PATH });
    expect(subjectOf(client)).toBe('untouched');
  });

  it('matches addresses case-insensitively', async () => {
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: 'Pro\n' });
    await run(deps, { to: 'ME@Example.COM', subject: 'shouty', html: BODY_PATH });
    expect(subjectOf(client)).toBe('[Pro] shouty');
  });

  it('ignores a blank host-tag file rather than emitting empty brackets', async () => {
    const { deps, client } = makeDeps({ ...BODY, [HOST_TAG]: '   \n' });
    await run(deps, { to: 'me@example.com', subject: 'no brackets', html: BODY_PATH });
    expect(subjectOf(client)).toBe('no brackets');
  });
});
