# Changelog (fork: weirdapps/outlook-access)

This is the **fork** of `BikS2013/outlook-tool`. Upstream's CHANGELOG is at
<https://github.com/BikS2013/outlook-tool/blob/master/CHANGELOG.md>.

Fork-only features (not in upstream):
- `--since` / `--until` / `--all` / `--max` pagination on `list-mail`
- `download-sharepoint-link` command + `sharepoint-client`/`sharepoint-schema`
  modules for SharePoint Bearer-session capture and ReferenceAttachment fetching
- `OutlookClient.listMessagesInFolderAll` auto-paginating helper

---

## [1.5.0] — 2026-04-22 (fork)

Phase B3: complete signature handling across send-mail / reply / forward,
including inline-image asset support so the user's NBG logo renders
correctly in delivered mail. Also strengthens CC-self compliance.

### Added
- **`send-mail` signature injection** — defaults to reading
  `~/.outlook-cli/signature.html` and appending after the user's HTML
  body (before `</body>`). New flags: `--signature <file>` (override),
  `--no-signature` (suppress). Plain-text bodies skip signature (HTML only).
- **Inline-image assets for signatures** — new util
  `src/util/signature-assets.ts`:
  - `extractCidReferences(html)` — finds `<img src="cid:XXX">` references.
  - `saveSignatureAssets(...)` + `loadManifest(...)` — persists/reads
    inline image bytes + manifest under `~/.outlook-cli/signature-assets/`.
  - `loadSignatureAttachments(...)` — for each cid in signature, returns a
    `SendFileAttachment` with `IsInline:true` + matching `ContentId`.
- **`capture-signature` extension** — after extracting signature HTML,
  scans for cid: refs and downloads matching attachments from the source
  message via the new `OutlookClient.listMessageAttachments`. Saves to
  `signature-assets/` with manifest. Result includes `assetsDir`,
  `inlineAssetCount`, `unmatchedCidRefs`.
- **`reply` / `reply-all` / `forward` inline-image support** — same
  signature loading pipeline; inline attachments POSTed via new
  `OutlookClient.addMessageAttachment` (M365 PATCH doesn't accept the
  Attachments collection — needs separate POST per attachment).
  Dedupe: if createReply preserved an attachment with matching cid,
  skip re-adding to avoid duplicates.
- **CC-self for reply / reply-all / forward** — new flag `--no-cc-self`
  (default ON, mirroring send-mail). Reply/reply-all merge into the
  server-populated CcRecipients; forward merges into user-supplied.
  ALWAYS adds self to CC unless already in CC (allows self in TO + CC,
  per CLAUDE.md compliance + audit-trail requirement).
- **`OutlookClient.listMessageAttachments(messageId)`** + **
  `addMessageAttachment(messageId, FileAttachment)`** — public methods
  used by capture-signature and reply.

### Changed
- `doRequest` / `executeFetch` / `buildHeaders` extended to accept PATCH
  (was only GET/POST).
- `SendMailOptions` types extended with `signature?` and `noSignature?`.
- `SendMailResult` and reply `ReplyResult` now report `signatureApplied`
  flag for caller verification.

### Notes
- Default behavior is "always cc self, always have signature, image
  displays properly" per user's explicit requirement.
- Stderr warning when signature has cid: refs but no matching asset
  (the image will display as broken in Outlook); user should re-run
  `outlook-cli capture-signature` to refresh the assets.
- Smoke verified end-to-end against live NBG mailbox: send-mail with
  unique-marker body → delivered with inline logo (verified bytes via
  sha256 prefix); reply with CC-self → present in CcRecipients; clean
  signature.html → no forwarded-thread cruft in delivered body.

### Known limitation
- `extractSignature` heuristic is best-effort — when capturing from a
  reply message, the "last-hr" fallback can grab the forwarded-thread
  block instead of the signature. User should hand-edit
  `~/.outlook-cli/signature.html` after capture to clean up if needed
  (or capture from a NEW outgoing mail rather than a reply). Improving
  the heuristic to detect `divRplyFwdMsg` markers is a follow-up.

---

## [1.4.0] — 2026-04-22 (fork)

Phase B2: reply / forward / signature. Together with v1.3.0 (B1), this
gives the fork a complete send pipeline that can replace the AppleScript
send path in `email-handler` (downstream migration tracked separately).

### Added
- `capture-signature [--from-message <id>] [--out <file>]` — extracts
  email signature from a SentItems message (or specified one). Heuristic
  priority: `<div id="Signature">` (Outlook web wrapper) → `<div class="elementToProof">`
  → last `<hr>` content → pre-reply-marker last `<p>` block → whole-body
  fallback. Saves to `~/.outlook-cli/signature.html` (mode 0600, parent
  dir 0700). Hand-edit to refine if heuristic captures too much.
- `reply <id>` — reply DRAFT via M365 `/me/messages/{id}/createReply`.
  Server pre-populates ToRecipients (original sender) and Subject ("RE:..."),
  auto-quotes the original body. Command then PATCHes the body to inject
  the user's `--html`/`--text` content + signature ABOVE the auto-quote.
  `--no-signature` to skip; `--signature <file>` to override default path.
- `reply-all <id>` — same pipeline using `/createReplyAll`. Server
  populates To+Cc with all original parties.
- `forward <id> --to <recipients>` — `/createForward`. ToRecipients is
  empty in the server response; command patches it from `--to` (also
  honors `--cc` and `--bcc`). Same draft-first default.
- All four new commands honor `--send-now` (skip draft, dispatch via
  sendDraft), `--no-open` (skip Outlook activation), `--dry-run`.
- New `OutlookClient` methods: `getMessage`, `updateMessage` (PATCH),
  `createReply`, `createReplyAll`, `createForward`. `doRequest` /
  `executeFetch` / `buildHeaders` extended to support PATCH method.
- New types: `CreateReplyResult`, `UpdateMessagePatch`, `GetMessageResult`,
  `GetMessageOptions`.

### Notes
- Inline `cid:` images and SharePoint reference attachments are still
  deferred (would have been B2 stretch — likely a v1.5.0 follow-up if
  needed).
- Smoke-verified against live NBG mailbox: capture-signature pulled the
  user's standard NBG signature; reply against a real message produced
  draft with auto-quote + signature + user content; forward pre-populated
  recipient correctly; missing `--to` on forward → BAD_USAGE.

---

## [1.3.0] — 2026-04-22 (fork)

Phase B1: send-mail core. Replaces the AppleScript send path for direct
CLI use; downstream `email-handler` migration tracked separately.

### Added
- `send-mail` command — new email composition with **draft-first default**.
  - `--to/--cc/--bcc` accept comma-separated strings AND/OR repeated flag.
  - `--html <file>` and/or `--text <file>` body sources.
  - `--attach <file>` repeatable (combined cap 30 MB).
  - `--cc-self` ON by default (resolves to authenticated UPN); `--no-cc-self` opt-out.
  - `--save-sent` ON by default (only meaningful with `--send-now`).
  - `--send-now` bypasses draft, POSTs directly to `/me/sendmail`.
  - `--no-open` suppresses Outlook desktop activation after draft.
  - `--dry-run` prints the would-send payload without contacting M365.
- `OutlookClient.sendMail(payload, opts?)` — immediate send via `/me/sendmail`.
- `OutlookClient.createDraft(payload)` → `{Id, WebLink, ConversationId}` via `/me/messages`.
- `OutlookClient.sendDraft(messageId)` — POST `/me/messages/{id}/send`.
- `src/util/open-outlook.ts` `activateOutlookApp()` — wraps macOS `open -a "Microsoft Outlook"`. Non-darwin = no-op.
- `redactMessageBodies()` extension — message Body.Content / HtmlBody / TextBody redacted from echoed-back error JSON (defense-in-depth; subject and ContentType preserved).

### Notes
- Inline `cid:` images, SharePoint reference attachments, and reply / reply-all / forward commands are deferred to **B2** (next plan).
- `email-handler` `/send-mail` skill still uses AppleScript path; migration to `outlook-cli send-mail` in a separate downstream PR.
- Smoke verified against live NBG mailbox: dry-run, draft creation, draft visible in Drafts folder, immediate send, attachment round-trip, Greek text preserved, error paths.

---

## [1.2.0] — 2026-04-22 (fork)

Cherry-pick from upstream `BikS2013/outlook-tool` v1.2.0 + v1.3.0
(commit `cca2f50`). Preserves all fork-only features above.

### Added
- `get-thread <id>` command — retrieves every message in a conversation by
  message id (or `conv:<conversationId>` to skip the resolve hop). Flags:
  `--body html|text|none` (default text), `--order asc|desc` (default asc).
- `OutlookClient.listMessagesByConversation(conversationId, opts?)` — public
  API method behind `get-thread`.
- `OutlookClient.countMessagesInFolder(folderId, opts?)` — server-side count
  via `$count=true`. Returns `{count, exact}`; `exact: false` when the
  server did not return `@odata.count`.
- `list-mail --from <iso|keyword>` and `--to <iso|keyword>` — date filters
  using ISO-8601 or `now` / `now + Nd` / `now - Nd` grammar.
- `list-mail --just-count` — returns only `{count, exact}` via server-side
  `$count=true`. O(1) cost regardless of mailbox size.
- `src/util/dates.ts` `parseTimestamp(raw)` — shared timestamp parser.
- `list-calendar --from/--to` — now accepts `now - Nd` keyword variant
  (inherited from shared `parseTimestamp`).
- `MessageSummary.ConversationId?: string` — optional field for `$select`.
- `ODataListResponse['@odata.count']?: number` — envelope field for count.

### Changed
- `list-mail --top` cap raised from 100 → 1000 (default still 10).
- `list-mail` rejects combining `--since/--until` with `--from/--to`
  (mutually exclusive — use `--from/--to` for new code).
- `list-mail` rejects combining `--just-count` with `--all` (count is one
  HTTP call by design; `--all` paginates).

### Deferred (downstream callers to migrate next)
- `outlook-bridge` MCP wrappers in `~/SourceCode/communications-marketplace`:
  add `outlook_get_thread` and `outlook_count_mail` tool wrappers; switch
  date filters from `--since` to `--from` (semantically equivalent).
- `second-brain` `outlook_export.py` in `~/SourceCode/second-brain`: switch
  `--since` to `--from` (no behavioral change).

---

## [1.1.1] — 2026-04-21 (fork)

### Fixed
- `outlook-cli` Node stdout/stderr pipe truncation when piping to `head`/
  `tail` etc. Caused by `process.exit()` not draining pipes. Added
  `exitWithDrain()` helper that `process.stdout.write('', cb)` before
  exiting. (PR #2)

## [1.1.0] — 2026-04-21 (fork)

### Added
- `list-mail --since <iso>` / `--until <iso>` — ISO-8601 date filters.
- `list-mail --all` / `--max <N>` — auto-paginate via `@odata.nextLink`
  with safety cap (default 10000, max 100000).
- `download-sharepoint-link <url>` command — fetches a SharePoint URL
  using the captured SharePoint Bearer session.
- `OutlookClient.listMessagesInFolderAll(folderId, opts, maxResults)` —
  auto-paginating helper backing `--all`.
- `src/http/sharepoint-client.ts`, `src/session/sharepoint-schema.ts`,
  `src/http/filter-builder.ts` — supporting modules.
- `outlook-cli login --sharepoint-host <host>` — extends interactive login
  to capture a separate Bearer session for SharePoint.
- (PR #1)
