# `outlook-cli send-mail` design spec — DECISIONS RESOLVED

> **Status:** RESOLVED 2026-04-22. Promoted to two implementation plans:
> - **B1 (core send):** `docs/superpowers/plans/2026-04-22-send-mail-b1-core.md`
> - **B2 (reply/forward/signature):** to be written after B1 ships
>
> Decision log (final):
> - **D1 Body format:** B — `--html <file>` and/or `--text <file>` accepted
> - **D2 Attachments:** C — file paths + inline `cid:` + SharePoint refs (B2)
> - **D3 Recipients:** A primary + B fallback (comma string + repeatable `--to`)
> - **D4 CC-self:** A — `--cc-self` ON by default, `--no-cc-self` opt-out
> - **D5 Send vs draft:** **Draft-first** — creates draft, activates Outlook desktop via `open -a "Microsoft Outlook"`, returns `{id, webLink}`; `--send-now` to bypass draft and send immediately; `--no-open` to suppress app activation
> - **D6 Reply/forward:** A — separate `reply`, `reply-all`, `forward` commands using M365 `/createReply` endpoints (B2)
> - **D7 Save to Sent:** A — always save, `--no-save-sent` opt-out
> - **D8 Styling:** A — no injection, CLI is transport. RTF doesn't apply to API send (`Body.ContentType` is HTML or Text only). Smoke verification by sending CC-to-self.
>
> Q1 Signature: bootstrap via `outlook-cli capture-signature [--from-message <id>]` (B2), saves to `~/.outlook-cli/signature.html`. Replies append it automatically; `--signature <path>` overrides; `--no-signature` skips.
>
> Q2 Draft review: Outlook desktop. CLI runs `open -a "Microsoft Outlook"` after draft creation to activate the app. User navigates to Drafts folder manually (no documented URL scheme to jump to a specific draft on Outlook for Mac).

**Goal:** Add a `send-mail` subcommand (and supporting `OutlookClient.sendMail()`) to `outlook-cli`, replacing the AppleScript-based send path used by `email-handler` (`/send-mail`, `/mail-review` reply drafts) and by the user's `pbcopy`-to-Outlook workflow.

**Why now:** Phase A (cherry-pick of upstream v1.2.0+v1.3.0) is done. The fork is the only place these features can land — upstream BikS keeps `outlook-cli` read-only by design. Our use case requires send (compliance: every outbound mail CCs `dimitrios.plessas@nbg.gr`; reply automation needs send capability without GUI Outlook running).

**Architecture (assumed; revisit if decisions invalidate it):**
- New file `src/commands/send-mail.ts` — input parsing, body building, attachment loading, dispatch.
- New method `OutlookClient.sendMail(payload)` — POSTs to `/api/v2.0/me/sendmail` (immediate send) or `/api/v2.0/me/messages` then `/send` (draft-then-send) using the same Bearer + auto-reauth machinery as reads.
- Reuses error taxonomy (`UsageError` exit 2, `UpstreamError` exit 5, `AuthError` exit 4).
- Adds a body-redaction guard so message body content never appears in error stderr (same principle as the existing token-redaction guard).

---

## Open decisions

Each decision below has a **Recommendation** (my default if you don't push back) and a **Trade-offs** section. Mark each with `[X]` next to your chosen option.

### Decision 1: Body format — HTML, text, or both?

- [ ] **A.** HTML only (matches your CLAUDE.md preference for Outlook compatibility).
- [ ] **B.** Both — accept `--html <file>` OR `--text <file>` OR both (multipart alternative).
- [ ] **C.** Auto-detect from file extension (`.html` → HTML, `.txt`/`.md` → text).

**Recommendation:** B. Your CLAUDE.md mandates HTML for everything that goes to Outlook recipients, but `email-handler`'s draft pipeline produces HTML directly so the CLI just needs to accept it. Supporting plain-text too costs ~10 LOC and unblocks scripted notifications (e.g., system alerts) where HTML overhead is wasteful.

**Trade-offs:** A is simplest. B is most flexible. C is "magic" and surprises users who name an HTML file `output.txt`.

### Decision 2: Attachments — file-path inputs only, or also inline images?

- [ ] **A.** File paths only (`--attach /path/to/file.pdf`, repeatable). Sent as base64-encoded `FileAttachment`.
- [ ] **B.** A + inline images via `--inline <cid>=<path>` (referenced from HTML body as `<img src="cid:logo">`).
- [ ] **C.** A + B + ReferenceAttachment (SharePoint URLs, no upload).

**Recommendation:** A initially. Adding B requires HTML parsing to inject `cid:` references and is rarely needed (most NBG comms use external image URLs anyway). C is interesting for SharePoint-stored documents but the user's send pipeline rarely deals with this — if you do, attach the URL as a hyperlink in the body.

**Trade-offs:** A keeps the surface tight (~50 LOC for attachment handling). B adds ~40 LOC + inline parser. C adds ~80 LOC and depends on SharePoint context (reuses our existing SharePoint session).

**Size limit:** M365 caps a single attachment at ~150 MB but `/sendmail` JSON body is capped much lower. Practical limit ~30 MB combined; we should validate at the CLI and return a clear `UsageError` rather than letting M365 reject with a cryptic message.

### Decision 3: Recipient syntax

Three styles to choose from:

- [ ] **A.** Comma-separated string per flag: `--to "alice@x.com, bob@y.com" --cc "carol@z.com"`
- [ ] **B.** Repeatable flag: `--to alice@x.com --to bob@y.com --cc carol@z.com`
- [ ] **C.** JSON array: `--to '["alice@x.com","bob@y.com"]'` (machine-friendly)
- [ ] **D.** All three accepted (commander parses, normalizes internally)

**Recommendation:** A primarily — matches the visual mental model of email To/CC fields. Keep B as fallback for shells where comma quoting is awkward. Skip C (JSON in shell args is painful).

**Trade-offs:** A is human-friendly. B is composable in scripts (`for r in $RECIPS; do ... --to "$r"; done`). D is the "yes to everything" trap that doubles parser complexity.

### Decision 4: CC-self default

Your CLAUDE.md says **"primary email: dimitrios.plessas@nbg.gr (always CC self)"**. Should the CLI auto-CC?

- [ ] **A.** Bake `--cc-self` flag into CLI; default ON; can disable with `--no-cc-self`. The CLI knows the authenticated UPN from the session (`session.account.upn`).
- [ ] **B.** Bake `--cc-self` flag; default OFF. Caller decides.
- [ ] **C.** Don't add the flag. The caller (`email-handler`) handles CC-self.

**Recommendation:** A. The compliance rule is global, not per-recipient. Forgetting CC-self once is the kind of mistake that gets you flagged in audits. Defaulting ON with an explicit override is the right safety posture for an internal tool you use directly.

**Trade-offs:** A is "safe by default" but hides the behavior from scripts that may not want it. C is "explicit at every call site" but invites mistakes.

### Decision 5: Send vs draft

- [ ] **A.** `send-mail` immediate send only. Caller can preview with `--dry-run`.
- [ ] **B.** `send-mail` immediate + separate `create-draft` command that outputs the draft id (caller can later POST to `/send` via a third command).
- [ ] **C.** `send-mail --draft` flag (no separate command).

**Recommendation:** A + `--dry-run`. Drafts add storage churn (every preview becomes a Draft to clean up later) and mostly serve human-review workflows that the user doesn't need from the CLI. `--dry-run` prints the JSON payload without POSTing — sufficient for testing.

**Trade-offs:** A is minimal. B is "complete" but needs a `send-draft <id>` command and adds 2 commands for one flow. C is simpler than B but `--draft` flag semantics aren't obvious (does it leave a draft? send the draft?).

### Decision 6: Reply / Forward — separate commands or flags?

- [ ] **A.** Separate commands: `reply <messageId>`, `reply-all <messageId>`, `forward <messageId>`. Each takes body + recipient flags as overrides.
- [ ] **B.** Flags on `send-mail`: `--reply-to <id>`, `--reply-all <id>`, `--forward <id>` (mutually exclusive).
- [ ] **C.** Skip entirely. `email-handler` plugin builds the body itself with the quoted-text and just calls `send-mail`.

**Recommendation:** C for now, A later if needed. M365 has dedicated `/me/messages/{id}/createReply`, `/createReplyAll`, `/createForward` endpoints that auto-quote the original — using them would require copying the auto-quoted body into the new message before sending, which is a different flow from `send-mail`. The `email-handler` plugin already builds reply bodies (with the user's signature, style guide, threading) and just needs send. Reply-as-a-CLI-feature can wait until there's a clear caller.

**Trade-offs:** A is the "complete" answer but doubles command surface and maintenance. B is compact but mixes concerns. C punts the decision but keeps the CLI tight.

### Decision 7: Save to Sent folder

M365 default: every sent message lands in `SentItems`. The `/me/sendmail` endpoint accepts `SaveToSentItems: true|false`.

- [ ] **A.** Always save to Sent (default). `--no-save-sent` override.
- [ ] **B.** Default save; flag `--save-sent` defaults to true; `--save-sent=false` to disable.
- [ ] **C.** Don't expose; always save (M365 default).

**Recommendation:** A. Audit/compliance benefits from always saving. The override exists for ephemeral scripted notifications (status pings, etc.) where Sent folder pollution is annoying.

**Trade-offs:** All three behave the same in the common case. A and B differ only in flag spelling (commander.js convention is `--no-X` for boolean negation, so A is idiomatic).

### Decision 8: Body styling / signature handling

Your CLAUDE.md has style rules: Aptos 12pt, color #404040, no `<p>` tags, lowercase subjects, etc. Should the CLI enforce/inject any of this?

- [ ] **A.** No styling injection. CLI accepts arbitrary HTML; caller is responsible for the style. (`email-handler` already produces conformant HTML.)
- [ ] **B.** Inject a global `<style>` block applying Aptos 12pt #404040 to body if not already present.
- [ ] **C.** Inject + lowercase the subject if it starts with an uppercase letter that isn't a Greek greeting.

**Recommendation:** A. Style is a per-recipient/per-message decision that the email-handler plugin is already specialized for. Baking it into the CLI removes flexibility (e.g., system alerts shouldn't have personal styling). The CLI's job is **transport**, not composition.

**Trade-offs:** A is the correct separation of concerns. B reduces "forgot to style" mistakes for direct CLI use but couples transport to presentation. C combines B with subject normalization and is way too magical for a low-level tool.

---

## Spec freeze checklist (after decisions)

Once Decisions 1-8 are made, the spec promotes to a plan with these tasks:

1. **Tests for `OutlookClient.sendMail()`** — fetch mock, payload shape verification, error mapping (HTTP 401/403/413/429 → CLI errors).
2. **`OutlookClient.sendMail()` implementation** — JSON body builder (depends on D1, D2, D3), reuse `doRequest` + `withAutoReauth`.
3. **Body-redaction extension** — current code redacts Bearer tokens and cookies from error stderr; add Body/HtmlBody field redaction.
4. **`src/commands/send-mail.ts`** — input parsing (depends on D3, D7), `--dry-run` (D5), CC-self resolution from session UPN (D4).
5. **CLI registration in `src/cli.ts`** — flags, help text.
6. **Smoke tests against live mailbox** — send to self, verify Sent folder entry (depends on D7), verify CC-self appears in headers (D4).
7. **CHANGELOG entry** — `[1.3.0] (fork)` adding `send-mail`.
8. **Downstream migration** (separate Plan):
   - `outlook-bridge` MCP — add `outlook_send_mail` tool wrapper.
   - `email-handler` `/send-mail` skill — switch from AppleScript to `outlook-cli send-mail` behind feature flag (rollout pattern matching the Phase A read-side migration).
   - Eventually retire AppleScript send path entirely → CLAUDE.md "Mail send: STILL osascript" rule gets removed.

## Estimated size

| Component | LOC (with tests) |
|---|---|
| `OutlookClient.sendMail` + types | ~120 |
| `src/commands/send-mail.ts` | ~150 |
| `src/cli.ts` registration | ~40 |
| Body-redaction extension | ~30 |
| Tests (mocks) | ~250 |
| Smoke tests (manual, doc only) | — |
| **Total** | **~590 LOC** |

Comparable to Phase A's `get-thread` + `--just-count` work (~600 LOC merged). Single sitting if decisions are clear.

## Risks / open technical questions

- **Throttling**: M365 send rate limit is 30 messages/min for personal mailboxes; mailbox-level limit ~10000/day. Bulk-send loops via this CLI risk hitting it. **Mitigation:** document the limit; do not add automatic backoff in v1.
- **DLP / compliance scanners**: NBG tenant likely has DLP rules that scan outbound mail. The CLI submits via the same path Outlook web does, so DLP applies the same way — no special handling needed.
- **Test mailbox**: smoke testing send means actually sending mail. **Mitigation:** all smokes send to self (`--to dimitrios.plessas@nbg.gr`); never to external addresses.
- **Idempotency**: if the network drops mid-POST and we retry, M365 may double-send. **Mitigation:** `/sendmail` is not idempotent by spec. We do NOT retry on send failures — let the caller decide. Auto-reauth on 401 is fine because that's a pre-send failure.
- **Greek text**: M365 handles UTF-8 natively. Verified end-to-end during Phase A smoke (Greek subjects + bodies preserved). Re-verify in send smoke.
