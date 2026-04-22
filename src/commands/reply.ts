// src/commands/reply.ts
//
// Reply / reply-all / forward commands. All three share the same composition
// pipeline:
//   1. POST /me/messages/{id}/createX → returns DRAFT with auto-quoted body
//   2. Read user's --html or --text body file
//   3. Append signature (default: ~/.outlook-cli/signature.html) unless --no-signature
//   4. Compose final body: USER_CONTENT + SIGNATURE + AUTO_QUOTED_ORIGINAL
//   5. PATCH /me/messages/{newDraftId} with the updated Body
//   6. For forward: also patch ToRecipients
//   7. Default: activate Outlook desktop. --send-now: also POST .../send.
//
// Mirrors the send-mail.ts shape (deps, options, draft-first, --no-open, etc).

import { promises as fs } from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import type {
  CreateReplyResult,
  OutlookClient,
  SendBody,
  SendBodyContentType,
  SendEmailAddress,
  SendFileAttachment,
  UpdateMessagePatch,
} from '../http/outlook-client';
import type { SessionFile } from '../session/schema';
import { activateOutlookApp } from '../util/open-outlook';
import { loadSignatureAttachments } from '../util/signature-assets';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export type ReplyKind = 'reply' | 'reply-all' | 'forward';

export interface ReplyDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
  activateOutlook?: () => Promise<void>;
  readFile?: (p: string) => Promise<Buffer>;
  homeDir?: () => string;
}

export interface ReplyOptions {
  /** Path to HTML body file (the user's NEW reply content). */
  html?: string;
  /** Or plain-text body file. */
  text?: string;
  /** Override signature file. Default: ~/.outlook-cli/signature.html */
  signature?: string;
  /** Suppress signature appending. */
  noSignature?: boolean;
  /** Forward only — additional TO recipients (comma string and/or repeat). */
  to?: string | string[];
  /** Forward only — additional CC recipients. */
  cc?: string | string[];
  /** Forward only — additional BCC recipients. */
  bcc?: string | string[];
  /**
   * When false, do NOT auto-CC the authenticated user. Default true
   * (matches send-mail's CLAUDE.md compliance default).
   */
  ccSelf?: boolean;
  /** Send immediately after composing (default: leave as draft + activate Outlook). */
  sendNow?: boolean;
  /** Activate Outlook desktop after draft (default true). */
  open?: boolean;
  /** Print payload-equivalent JSON, do not contact M365. */
  dryRun?: boolean;
}

export interface ReplyResult {
  kind: ReplyKind;
  mode: 'draft' | 'sent' | 'dry-run';
  sourceMessageId: string;
  /** New draft id (or sent message id if --send-now). */
  id?: string;
  webLink?: string;
  conversationId?: string;
  subject: string;
  /** Whether the auto-quoted original was preserved. */
  hasQuotedOriginal: boolean;
  /** Whether a signature was appended. */
  signatureApplied: boolean;
  /** For forward only: the final To list (server pre-pop + user-supplied). */
  to: string[];
}

const DEFAULT_SIG_REL = path.join('.outlook-cli', 'signature.html');

export async function run(
  deps: ReplyDeps,
  kind: ReplyKind,
  sourceMessageId: string,
  opts: ReplyOptions = {},
): Promise<ReplyResult> {
  if (typeof sourceMessageId !== 'string' || sourceMessageId.length === 0) {
    throw new UsageError(`${kind}: message id argument is required`);
  }

  const hasHtml = typeof opts.html === 'string' && opts.html.length > 0;
  const hasText = typeof opts.text === 'string' && opts.text.length > 0;
  if (!hasHtml && !hasText) {
    throw new UsageError(
      `${kind}: at least one of --html <file> or --text <file> is required`,
    );
  }

  // -------- User body --------
  const reader = deps.readFile ?? ((p: string) => fs.readFile(p));
  let userBodyContentType: SendBodyContentType;
  let userBody: string;
  if (hasHtml) {
    userBodyContentType = 'HTML';
    userBody = await readBodyFile(reader, opts.html as string, '--html', kind);
  } else {
    userBodyContentType = 'Text';
    userBody = await readBodyFile(reader, opts.text as string, '--text', kind);
    // Wrap text body in <p> so we can splice it into the HTML quoted draft.
    userBody = `<p>${escapeHtml(userBody).replace(/\n/g, '<br>')}</p>`;
  }

  // -------- Signature --------
  const home = (deps.homeDir ?? os.homedir)();
  const sigPath = opts.signature ?? path.join(home, DEFAULT_SIG_REL);
  let signatureHtml = '';
  let signatureApplied = false;
  let signatureInlineAttachments: SendFileAttachment[] = [];
  if (opts.noSignature !== true) {
    try {
      const buf = await reader(sigPath);
      signatureHtml = buf.toString('utf-8');
      signatureApplied = signatureHtml.length > 0;
      if (signatureApplied) {
        // Load matching inline images (cid: refs in signature → assets dir).
        const assetsDir = path.join(path.dirname(sigPath), 'signature-assets');
        const sigAssets = await loadSignatureAttachments({
          signatureHtml,
          assetsDir,
          reader,
        });
        signatureInlineAttachments = sigAssets.attachments;
        if (sigAssets.unmatchedRefs.length > 0) {
          process.stderr.write(
            `${kind}: warning — ${sigAssets.unmatchedRefs.length} cid: reference(s) ` +
              `in signature have no matching inline image asset (will display as ` +
              `broken image): ${sigAssets.unmatchedRefs.join(', ')}. ` +
              `Re-run \`outlook-cli capture-signature\` to refresh.\n`,
          );
        }
      }
    } catch {
      // Signature file missing or unreadable — note in result, don't fail.
      signatureApplied = false;
    }
  }

  // -------- Forward-only recipient parsing --------
  const userTo = parseRecipients(opts.to);
  const userCc = parseRecipients(opts.cc);
  const userBcc = parseRecipients(opts.bcc);

  if (kind !== 'forward' && (userTo.length > 0 || userCc.length > 0 || userBcc.length > 0)) {
    throw new UsageError(
      `${kind}: --to/--cc/--bcc are only meaningful for forward (use the original ` +
        'message thread participants — this command is reply/reply-all).',
    );
  }
  if (kind === 'forward' && userTo.length === 0) {
    throw new UsageError('forward: --to <recipients> is required');
  }

  const session = await ensureSession(deps);

  // -------- Dry-run --------
  if (opts.dryRun === true) {
    return {
      kind,
      mode: 'dry-run',
      sourceMessageId,
      subject: '(unknown — would call createReply server-side)',
      hasQuotedOriginal: true,
      signatureApplied,
      to: kind === 'forward' ? userTo : [],
    };
  }

  // -------- Server-side draft creation --------
  const client = deps.createClient(session);
  let draft: CreateReplyResult;
  try {
    if (kind === 'reply') {
      draft = await client.createReply(sourceMessageId);
    } else if (kind === 'reply-all') {
      draft = await client.createReplyAll(sourceMessageId);
    } else {
      draft = await client.createForward(sourceMessageId);
    }
  } catch (err) {
    throw mapHttpError(err);
  }

  // -------- Compose final body --------
  // Strategy: M365 returned an HTML draft with the original auto-quoted at the
  // bottom (typically wrapped in <hr> + <div>). We INSERT the user's content
  // (and signature) at the TOP of the body — above the quoted region.
  const composed = composeReplyBody(draft.Body.Content, userBody, signatureHtml);

  // -------- Patch the draft --------
  const patch: UpdateMessagePatch = {
    Body: { ContentType: 'HTML', Content: composed },
  };
  if (kind === 'forward') {
    const merged = userTo.map((addr) => ({ EmailAddress: { Address: addr } }));
    patch.ToRecipients = merged;
    if (userCc.length > 0) {
      patch.CcRecipients = userCc.map((addr) => ({ EmailAddress: { Address: addr } }));
    }
    if (userBcc.length > 0) {
      patch.BccRecipients = userBcc.map((addr) => ({ EmailAddress: { Address: addr } }));
    }
  }

  // -------- CC-self injection (compliance default: ON) --------
  // ALWAYS add self to CC unless already in CC (audit/archive — user's
  // CLAUDE.md mandates this, and they explicitly want self-CC even when
  // they're also the To recipient e.g., when replying to own mail).
  // For reply/reply-all: merge into the server-populated CcRecipients.
  // For forward: merge into the user-supplied CcRecipients (already in patch).
  // Suppress only if --no-cc-self.
  if (opts.ccSelf !== false) {
    const selfUpn = session.account?.upn;
    if (typeof selfUpn === 'string' && selfUpn.length > 0) {
      const existingCc = kind === 'forward'
        ? (patch.CcRecipients ?? [])
        : (draft.CcRecipients ?? []);
      const lower = selfUpn.toLowerCase();
      const alreadyInCc = existingCc.some(
        (r) => r.EmailAddress.Address.toLowerCase() === lower,
      );
      if (alreadyInCc) {
        // Preserve existing CC list in the patch so other patched fields
        // don't accidentally clear it (PATCH semantics: omitted = unchanged,
        // present = replace).
        patch.CcRecipients = existingCc;
      } else {
        patch.CcRecipients = [
          ...existingCc,
          { EmailAddress: { Address: selfUpn } },
        ];
      }
    }
  }
  try {
    await client.updateMessage(draft.Id, patch);
  } catch (err) {
    throw mapHttpError(err);
  }

  // -------- Splice signature inline attachments --------
  // Each inline image must be POSTed to /messages/{id}/attachments
  // separately (PATCH doesn't accept the Attachments collection).
  //
  // Dedupe: M365's createReply / createForward preserves the original
  // message's attachments. If the original was a previous reply that
  // already contained our signature image, the cid is already in the
  // draft. List existing attachments and skip any cid already present.
  if (signatureInlineAttachments.length > 0) {
    let existingCids = new Set<string>();
    try {
      const existing = await client.listMessageAttachments(draft.Id);
      for (const a of existing) {
        if (typeof a.ContentId === 'string' && a.ContentId.length > 0) {
          existingCids.add(a.ContentId);
        }
      }
    } catch {
      // Non-fatal: if list fails, attempt to add all and let M365 dedupe.
    }
    for (const inlineAtt of signatureInlineAttachments) {
      if (inlineAtt.ContentId && existingCids.has(inlineAtt.ContentId)) {
        continue; // already in draft (preserved from original)
      }
      try {
        await client.addMessageAttachment(draft.Id, inlineAtt);
      } catch (err) {
        // Non-fatal: the draft is created and the body refers to the cid;
        // the image will just show as broken. Warn once per failure.
        const msg = err instanceof Error ? err.message : String(err);
        process.stderr.write(
          `${kind}: failed to attach inline image "${inlineAtt.Name}" (${inlineAtt.ContentId}): ${msg}\n`,
        );
      }
    }
  }

  // -------- Send-now or activate Outlook --------
  if (opts.sendNow === true) {
    try {
      await client.sendDraft(draft.Id);
    } catch (err) {
      throw mapHttpError(err);
    }
    return {
      kind,
      mode: 'sent',
      sourceMessageId,
      id: draft.Id,
      webLink: draft.WebLink,
      conversationId: draft.ConversationId,
      subject: draft.Subject,
      hasQuotedOriginal: true,
      signatureApplied,
      to: kind === 'forward' ? userTo : draft.ToRecipients.map((r) => r.EmailAddress.Address),
    };
  }

  if (opts.open !== false) {
    const activate = deps.activateOutlook ?? activateOutlookApp;
    try {
      await activate();
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      process.stderr.write(`${kind}: Outlook activation failed: ${msg}\n`);
    }
  }

  return {
    kind,
    mode: 'draft',
    sourceMessageId,
    id: draft.Id,
    webLink: draft.WebLink,
    conversationId: draft.ConversationId,
    subject: draft.Subject,
    hasQuotedOriginal: true,
    signatureApplied,
    to: kind === 'forward' ? userTo : draft.ToRecipients.map((r) => r.EmailAddress.Address),
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Insert user body (and signature) at the TOP of the auto-quoted draft body
 * returned by createReply/All/Forward.
 *
 * Strategy: find the first <body...> tag (M365 wraps the auto-quote in a
 * full HTML doc including <html><head>...<body>) and inject our content
 * immediately after it. If there's no <body> tag, prepend to the whole HTML.
 */
export function composeReplyBody(
  quotedDraftHtml: string,
  userBodyHtml: string,
  signatureHtml: string,
): string {
  const sigBlock = signatureHtml.length > 0
    ? `\n<br><br>${signatureHtml}`
    : '';
  const userBlock = `${userBodyHtml}${sigBlock}\n<br>`;

  // Look for <body ...> tag
  const bodyMatch = quotedDraftHtml.match(/<body\b[^>]*>/i);
  if (bodyMatch && typeof bodyMatch.index === 'number') {
    const bodyTagEnd = bodyMatch.index + bodyMatch[0].length;
    return (
      quotedDraftHtml.slice(0, bodyTagEnd) +
      userBlock +
      quotedDraftHtml.slice(bodyTagEnd)
    );
  }

  // No <body> tag — prepend.
  return userBlock + quotedDraftHtml;
}

function parseRecipients(input: string | string[] | undefined): string[] {
  if (input === undefined) return [];
  const raw = Array.isArray(input) ? input : [input];
  const out: string[] = [];
  for (const item of raw) {
    if (typeof item !== 'string') continue;
    for (const part of item.split(',')) {
      const trimmed = part.trim();
      if (trimmed.length === 0) continue;
      if (!trimmed.includes('@')) {
        throw new UsageError(`invalid recipient address (no '@'): ${trimmed}`);
      }
      out.push(trimmed);
    }
  }
  return out;
}

async function readBodyFile(
  reader: (p: string) => Promise<Buffer>,
  filePath: string,
  flagName: string,
  kind: ReplyKind,
): Promise<string> {
  try {
    const buf = await reader(filePath);
    return buf.toString('utf-8');
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new UsageError(
      `${kind}: ${flagName} file read failed (${filePath}): ${msg}`,
    );
  }
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Re-export so CLI + tests can import.
export { UsageError };

// Re-export type for CLI.
export type { SendBody, SendEmailAddress };
