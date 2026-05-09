# Plan B1: send-mail core (immediate + draft) implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `outlook-cli send-mail` command with draft-first default, replacing AppleScript send for the user's direct workflow and the foundation for `email-handler`'s migration.

**Scope (B1 only):** New command + 4 client methods + tests + smoke. Excludes reply/forward/signature (those are B2).

**Architecture:**

- New `OutlookClient` methods: `createDraft(payload)`, `sendDraft(id)`, `sendMail(payload)` (immediate). All POST to `/api/v2.0/me/...` using existing Bearer + auto-reauth.
- `src/commands/send-mail.ts` — input parsing, body building (HTML and/or text), attachment loading (file paths only in B1), recipient parsing (comma-string + repeatable), CC-self injection from session UPN, dispatch.
- `src/util/open-outlook.ts` — small helper that runs `open -a "Microsoft Outlook"` via `child_process.spawn`.
- Body redaction extension for error stderr (don't leak email body content into logs).

**Tech Stack:** TypeScript, vitest, commander.js, Node 20+, macOS `open` command.

---

## Baseline state (verified post-Phase-A 2026-04-22)

- Branch: `master` at `9c1cfb4` (v1.2.0, post-cherry-pick)
- Tests: **285 passing across 27 files**
- Globally linked: `outlook-cli` → v1.2.0 dist

## Target end-state (B1 only)

- Version `1.3.0`
- Tests: ~315 passing (285 + ~30 new)
- New CLI surface: `send-mail`
- New API: `OutlookClient.createDraft`, `sendDraft`, `sendMail`, `addAttachment`
- Body redaction guard extended to message body fields

## Files map

| File                                       | Action                         | LOC       |
| ------------------------------------------ | ------------------------------ | --------- |
| `src/http/outlook-client.ts`               | modify — add 4 methods + types | +200      |
| `src/http/errors.ts`                       | modify — extend redaction      | +20       |
| `src/commands/send-mail.ts`                | create                         | +250      |
| `src/util/open-outlook.ts`                 | create                         | +30       |
| `src/cli.ts`                               | modify — register command      | +60       |
| `test_scripts/outlook-client-send.spec.ts` | create                         | +250      |
| `test_scripts/commands-send-mail.spec.ts`  | create                         | +250      |
| `test_scripts/util-open-outlook.spec.ts`   | create                         | +50       |
| `package.json`                             | modify — version bump          | +1        |
| `CHANGELOG.md`                             | modify — add 1.3.0 entry       | +25       |
| **Total**                                  |                                | **~1136** |

---

## Task 1: Pre-flight branch + baseline

**Files:** none (branch + verify only)

- [ ] **Step 1.1:** Create feature branch

  ```bash
  cd ~/SourceCode/outlook-access
  git checkout -b feat/send-mail-b1-core
  ```

- [ ] **Step 1.2:** Verify baseline `npm test` → 285 passing

- [ ] **Step 1.3:** Verify global `outlook-cli --version` → 1.2.0

---

## Task 2: Util — `open-outlook.ts`

**Files:**

- Create: `src/util/open-outlook.ts`
- Test: `test_scripts/util-open-outlook.spec.ts`

- [ ] **Step 2.1:** Write failing test in `test_scripts/util-open-outlook.spec.ts`

  ```ts
  import { describe, it, expect, vi } from 'vitest';
  import * as child from 'child_process';
  import { activateOutlookApp } from '../src/util/open-outlook';

  describe('activateOutlookApp', () => {
    it('spawns `open -a "Microsoft Outlook"` and resolves on close 0', async () => {
      const spawnSpy = vi.spyOn(child, 'spawn').mockReturnValue({
        on: (ev: string, cb: (n: number) => void) => {
          if (ev === 'close') queueMicrotask(() => cb(0));
          return this;
        },
        unref: () => {},
      } as never);
      await expect(activateOutlookApp()).resolves.toBeUndefined();
      expect(spawnSpy).toHaveBeenCalledWith('open', ['-a', 'Microsoft Outlook'], {
        stdio: 'ignore',
        detached: true,
      });
    });

    it('rejects with code when open exits non-zero', async () => {
      vi.spyOn(child, 'spawn').mockReturnValue({
        on: (ev: string, cb: (n: number) => void) => {
          if (ev === 'close') queueMicrotask(() => cb(1));
          return this;
        },
        unref: () => {},
      } as never);
      await expect(activateOutlookApp()).rejects.toThrow(/exited with code 1/);
    });
  });
  ```

- [ ] **Step 2.2:** Run test → fails (file doesn't exist)

- [ ] **Step 2.3:** Implement

  ```ts
  // src/util/open-outlook.ts
  //
  // Activates Microsoft Outlook desktop app on macOS via `open -a`.
  // No-op on non-darwin (logs to stderr).
  import { spawn } from 'child_process';

  export async function activateOutlookApp(): Promise<void> {
    if (process.platform !== 'darwin') {
      process.stderr.write('open-outlook: skipping (non-darwin platform)\n');
      return;
    }
    return new Promise((resolve, reject) => {
      const child = spawn('open', ['-a', 'Microsoft Outlook'], { stdio: 'ignore', detached: true });
      child.on('close', (code) => {
        if (code === 0) resolve();
        else reject(new Error(`open -a "Microsoft Outlook" exited with code ${code}`));
      });
      child.on('error', reject);
      child.unref();
    });
  }
  ```

- [ ] **Step 2.4:** Run test → passes

- [ ] **Step 2.5:** Commit

  ```bash
  git add src/util/open-outlook.ts test_scripts/util-open-outlook.spec.ts
  git commit -m "feat(util): add activateOutlookApp helper"
  ```

---

## Task 3: Body redaction extension

**Files:**

- Modify: `src/http/errors.ts` (or wherever the existing token-redaction lives)

- [ ] **Step 3.1:** Locate redaction. `grep -n "redact\|REDACT\|Bearer\|cookie" src/http/errors.ts`

- [ ] **Step 3.2:** Extend redaction regex to catch `Body":{"Content":"..."` and `HtmlBody":"..."` patterns. Replace value with `[REDACTED-BODY]`.

- [ ] **Step 3.3:** Add tests — error with body POST URL gets `[REDACTED-BODY]` substitution.

- [ ] **Step 3.4:** Run tests → pass.

- [ ] **Step 3.5:** Commit

  ```bash
  git commit -m "feat(errors): redact Body/HtmlBody from error stderr"
  ```

---

## Task 4: `OutlookClient.sendMail` (immediate send)

**Files:**

- Modify: `src/http/outlook-client.ts` — types + interface + impl
- Test: `test_scripts/outlook-client-send.spec.ts`

- [ ] **Step 4.1:** Define types

  ```ts
  export type EmailAddress = { Address: string; Name?: string };
  export type BodyContent = { ContentType: 'HTML' | 'Text'; Content: string };
  export type FileAttachmentInput = {
    '@odata.type': '#Microsoft.OutlookServices.FileAttachment';
    Name: string;
    ContentType?: string;
    ContentBytes: string; // base64
    IsInline?: boolean;
    ContentId?: string;
  };
  export interface SendMailPayload {
    Subject: string;
    Body: BodyContent;
    ToRecipients: { EmailAddress: EmailAddress }[];
    CcRecipients?: { EmailAddress: EmailAddress }[];
    BccRecipients?: { EmailAddress: EmailAddress }[];
    Attachments?: FileAttachmentInput[];
  }
  export interface SendMailOptions {
    saveToSentItems?: boolean; // default true
  }
  ```

- [ ] **Step 4.2:** Add method to `OutlookClient` interface

  ```ts
  sendMail(payload: SendMailPayload, opts?: SendMailOptions): Promise<void>;
  ```

- [ ] **Step 4.3:** Write 5 failing tests in `test_scripts/outlook-client-send.spec.ts` covering:
  - POSTs to `/api/v2.0/me/sendmail` with `{Message: <payload>, SaveToSentItems: true}`
  - Honors `saveToSentItems: false`
  - Throws AuthError on 401 (no auto-reauth in tests via `noAutoReauth: true`)
  - Throws UpstreamError with redacted body on 400
  - Throws on 413 (payload too large) with hint about attachment limits

- [ ] **Step 4.4:** Implement method (mirrors existing `doPost` pattern)

  ```ts
  async function sendMail(payload: SendMailPayload, opts: SendMailOptions = {}): Promise<void> {
    const body = {
      Message: payload,
      SaveToSentItems: opts.saveToSentItems !== false,
    };
    try {
      await doPost('/api/v2.0/me/sendmail', body);
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }
  ```

- [ ] **Step 4.5:** Add to return statement at bottom of `createOutlookClient`

- [ ] **Step 4.6:** Tests pass; commit

  ```bash
  git commit -m "feat(client): add sendMail() — immediate POST /me/sendmail"
  ```

---

## Task 5: `OutlookClient.createDraft` + `sendDraft`

**Files:**

- Modify: `src/http/outlook-client.ts`
- Test: extend `test_scripts/outlook-client-send.spec.ts`

- [ ] **Step 5.1:** Add types

  ```ts
  export interface CreateDraftResult {
    Id: string;
    WebLink: string;
    ConversationId?: string;
  }
  ```

- [ ] **Step 5.2:** Add interface methods

  ```ts
  createDraft(payload: SendMailPayload): Promise<CreateDraftResult>;
  sendDraft(messageId: string): Promise<void>;
  ```

- [ ] **Step 5.3:** Write failing tests:
  - `createDraft` POSTs to `/api/v2.0/me/messages` with the message body, returns `{Id, WebLink, ConversationId}`
  - `sendDraft(id)` POSTs to `/api/v2.0/me/messages/{id}/send` with empty body
  - Both honor auth/reauth chain

- [ ] **Step 5.4:** Implement

  ```ts
  async function createDraft(payload: SendMailPayload): Promise<CreateDraftResult> {
    try {
      const resp = await doPost<
        SendMailPayload,
        { Id: string; WebLink: string; ConversationId?: string }
      >('/api/v2.0/me/messages', payload);
      return { Id: resp.Id, WebLink: resp.WebLink, ConversationId: resp.ConversationId };
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }

  async function sendDraft(messageId: string): Promise<void> {
    if (!messageId) throw new Error('outlook-client: sendDraft requires non-empty messageId');
    try {
      await doPost(`/api/v2.0/me/messages/${encodeURIComponent(messageId)}/send`, {});
    } catch (err) {
      throw mapHttpToCliError(err);
    }
  }
  ```

- [ ] **Step 5.5:** Wire to return statement; tests pass; commit

  ```bash
  git commit -m "feat(client): add createDraft + sendDraft for staged-send workflow"
  ```

---

## Task 6: `send-mail` command

**Files:**

- Create: `src/commands/send-mail.ts`
- Test: `test_scripts/commands-send-mail.spec.ts`

- [ ] **Step 6.1:** Define `SendMailDeps` and `SendMailOptions` types in the new file

- [ ] **Step 6.2:** Write 12 failing tests covering:
  - Required fields: `--to` and `--subject` (UsageError if missing)
  - At least one body source: `--html` or `--text` (UsageError if both missing)
  - Both `--html` and `--text`: HTML wins as primary, text becomes alternative (or merge into multipart — see implementation)
  - Recipient parsing: `--to "a@x.com, b@y.com"` → 2 recipients; multiple `--to` flags → all collected
  - Default CC-self: payload includes `CcRecipients` with `session.account.upn` unless `--no-cc-self`
  - Attachments: `--attach /path/to/file.pdf` reads file, base64-encodes, sets `Name`, `ContentType` (from extension via `mime-types` or hardcoded map), `IsInline: false`
  - Subject lowercase mention from CLAUDE.md is NOT auto-applied (user can pass any case; we don't normalize)
  - Default behavior: calls `client.createDraft(payload)`, returns `{id, webLink}`, calls `activateOutlookApp()`
  - `--send-now`: calls `client.sendMail(payload)`, no `activateOutlookApp` call
  - `--dry-run`: prints payload as JSON, no client calls
  - `--no-open`: skips `activateOutlookApp` even on draft path
  - Error: file doesn't exist for `--html`/`--text`/`--attach` → UsageError with file path

- [ ] **Step 6.3:** Implement `run()` function. ~250 LOC including:
  - File reads (sync, since this is short-lived CLI)
  - Recipient parser (`parseRecipients(input: string | string[]): EmailAddress[]`)
  - Attachment loader (`loadAttachment(path: string): FileAttachmentInput`)
  - Payload builder
  - Dispatch (dry-run / send-now / draft)

- [ ] **Step 6.4:** Tests pass; commit

  ```bash
  git commit -m "feat(send-mail): command implementation with draft-first default"
  ```

---

## Task 7: CLI registration

**Files:**

- Modify: `src/cli.ts`

- [ ] **Step 7.1:** Import `* as sendMail from './commands/send-mail'`

- [ ] **Step 7.2:** Register command after `move-mail`

  ```ts
  // -------- send-mail --------
  program
    .command('send-mail')
    .description('Send a new email (default: creates draft and activates Outlook desktop)')
    .requiredOption('--to <recipients...>', 'TO recipients (comma-separated or repeat flag)')
    .option('--cc <recipients...>', 'CC recipients')
    .option('--bcc <recipients...>', 'BCC recipients')
    .requiredOption('--subject <s>', 'Subject line')
    .option('--html <file>', 'HTML body file path')
    .option('--text <file>', 'Plain-text body file path')
    .option(
      '--attach <file>',
      'Attach file (repeatable)',
      (v: string, acc: string[] = []) => [...acc, v],
      [] as string[],
    )
    .option('--no-cc-self', 'Suppress automatic CC to authenticated user')
    .option('--no-save-sent', 'Do not save to Sent folder')
    .option('--send-now', 'Send immediately, skip draft', false)
    .option('--no-open', 'Do not activate Outlook desktop after draft creation')
    .option('--dry-run', 'Print payload JSON, do not send', false)
    .action(
      makeAction<SendMailCmdOpts, []>(program, async (deps, g, cmdOpts) => {
        const result = await sendMail.run(deps, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );
  ```

- [ ] **Step 7.3:** Run full `npm test` — all 285+30 should pass; commit

---

## Task 8: Smoke verification (manual, against live mailbox)

**Files:** none (smoke only)

- [ ] **Step 8.1:** Build

  ```bash
  npm run build
  ```

- [ ] **Step 8.2:** Self-test draft creation

  ```bash
  echo '<p>test draft from outlook-cli</p>' > /tmp/test-body.html
  outlook-cli --quiet send-mail \
    --to you@example.com \
    --subject "test draft from outlook-cli" \
    --html /tmp/test-body.html
  ```

  Expected:
  - `{id, webLink}` printed
  - Microsoft Outlook desktop activates (focuses)
  - Switch to Outlook → Drafts folder → see new draft "test draft from outlook-cli"
  - Verify CC-self present (your address in CC)

- [ ] **Step 8.3:** Self-test immediate send

  ```bash
  outlook-cli --quiet send-mail \
    --to you@example.com \
    --subject "test immediate send from outlook-cli" \
    --html /tmp/test-body.html \
    --send-now
  ```

  Expected:
  - Empty stdout (or `{}`)
  - Mail arrives in Inbox (CC-self) within seconds
  - Mail also in Sent folder
  - HTML renders correctly (paragraph tag becomes paragraph, not literal `<p>`)

- [ ] **Step 8.4:** Self-test attachment

  ```bash
  echo "test attachment content" > /tmp/test-attach.txt
  outlook-cli --quiet send-mail \
    --to you@example.com \
    --subject "test attachment from outlook-cli" \
    --html /tmp/test-body.html \
    --attach /tmp/test-attach.txt \
    --send-now
  ```

  Expected: mail arrives with `test-attach.txt` attached; downloads correctly.

- [ ] **Step 8.5:** Self-test error paths
  - `--to` missing → exit 2, BAD_USAGE
  - `--html` pointing to nonexistent file → exit 2, BAD_USAGE with file path
  - `--dry-run` → JSON payload printed, no mail sent

- [ ] **Step 8.6:** Greek text round-trip

  ```bash
  echo '<p>Καλημέρα — δοκιμή ελληνικών</p>' > /tmp/test-greek.html
  outlook-cli send-mail \
    --to you@example.com \
    --subject "δοκιμή ελληνικών from outlook-cli" \
    --html /tmp/test-greek.html \
    --send-now
  ```

  Expected: subject and body both display Greek correctly in Outlook.

---

## Task 9: Bump version + CHANGELOG

**Files:** `package.json`, `CHANGELOG.md`

- [ ] **Step 9.1:** `package.json` version → `1.3.0`

- [ ] **Step 9.2:** Prepend to CHANGELOG.md

  ```markdown
  ## [1.3.0] — 2026-04-22 (fork)

  Phase B1: send-mail core.

  ### Added

  - `send-mail` command — new email composition with draft-first default.
  - `OutlookClient.sendMail()` — immediate send via `/me/sendmail`.
  - `OutlookClient.createDraft()` + `sendDraft()` — staged-send via `/me/messages` + `/send`.
  - `src/util/open-outlook.ts` `activateOutlookApp()` — macOS `open -a` wrapper.
  - Body/HtmlBody redaction in error stderr (extends existing token redaction).

  ### Notes

  - Default workflow: creates draft, returns `{id, webLink}`, activates Outlook desktop.
  - `--send-now` bypasses draft.
  - `--cc-self` defaults ON; resolves to authenticated UPN from session.
  - Attachments: file paths only in B1 (inline `cid:` and SharePoint refs come in B2).
  - Reply/forward and `capture-signature` deferred to B2.
  ```

- [ ] **Step 9.3:** Commit

  ```bash
  git commit -m "chore: bump to 1.3.0 + CHANGELOG B1"
  ```

---

## Task 10: Push, PR, merge, relink

- [ ] **Step 10.1:** Push branch

  ```bash
  git push -u origin feat/send-mail-b1-core
  ```

- [ ] **Step 10.2:** Create PR — **explicitly target `weirdapps/outlook-access`** (don't repeat the upstream-PR mistake from Phase A)

  ```bash
  gh pr create --repo weirdapps/outlook-access --base master --head feat/send-mail-b1-core --title "feat(B1): add send-mail command (draft-first default)" --body "..."
  ```

- [ ] **Step 10.3:** Self-merge after smoke verified

  ```bash
  gh pr merge <PR#> --repo weirdapps/outlook-access --squash --delete-branch
  ```

- [ ] **Step 10.4:** Switch master, rebuild, verify

  ```bash
  git checkout master && git pull --ff-only && npm run build
  outlook-cli --version  # 1.3.0
  outlook-cli send-mail --help
  ```

---

## Rollback plan

If anything breaks:

```bash
git checkout master
git reset --hard 9c1cfb4   # back to v1.2.0
npm run build
outlook-cli --version  # 1.2.0
```

The `email-handler` plugin's `/send-mail` skill is unchanged in B1 (still uses AppleScript). It will start using `outlook-cli send-mail` only in a separate downstream migration PR after B2 lands.

---

## After B1 ships → B2 plan

`docs/superpowers/plans/2026-04-22-send-mail-b2-replies.md` (TBW after B1):

- `capture-signature` command + heuristic extraction
- `reply <id>` / `reply-all <id>` / `forward <id>` commands
- Inline `cid:` attachments (`--inline cid=path`)
- SharePoint reference attachments (`--ref-attach <url>`, uses our SharePoint session)
- Same draft-first default
- Bumps to 1.4.0
