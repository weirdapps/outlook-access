// src/commands/send-mail.ts
//
// Compose and send (or draft) a new email via M365 v2.0 REST.
//
// Default behavior: creates a DRAFT, returns {Id, WebLink}, and activates
// Microsoft Outlook desktop on macOS so the user can review/send manually.
// Pass `--send-now` to bypass the draft and POST to /me/sendmail directly.

import { promises as fs } from 'node:fs';
import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import type {
  CreateDraftResult,
  OutlookClient,
  SendBodyContentType,
  SendFileAttachment,
  SendMailPayload,
} from '../http/outlook-client';
import type { SessionFile } from '../session/schema';
import { activateOutlookApp } from '../util/open-outlook';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

// Re-export so callers (CLI + tests) can import from the command module.
export { UsageError };

export interface SendMailDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
  /** Optional override for tests — defaults to the real activateOutlookApp. */
  activateOutlook?: () => Promise<void>;
  /** Optional override for tests — defaults to fs.readFile. */
  readFile?: (p: string) => Promise<Buffer>;
}

export interface SendMailOptions {
  /** Repeated --to or comma-separated string. Required. */
  to?: string | string[];
  cc?: string | string[];
  bcc?: string | string[];
  /** Required. */
  subject?: string;
  /** Path to HTML body file. Either --html or --text required (or both). */
  html?: string;
  /** Path to plain-text body file. */
  text?: string;
  /** Repeatable --attach <file>. */
  attach?: string[];
  /** When false, do NOT auto-CC the authenticated user. Default true. */
  ccSelf?: boolean;
  /** When false, do not save to SentItems. Default true. */
  saveSent?: boolean;
  /** When true, send immediately (skip draft + Outlook activation). */
  sendNow?: boolean;
  /** When false, do not activate Outlook desktop after draft. Default true. */
  open?: boolean;
  /** When true, print the JSON payload and exit without contacting Outlook. */
  dryRun?: boolean;
}

export interface SendMailResult {
  mode: 'draft' | 'sent' | 'dry-run';
  id?: string;
  webLink?: string;
  conversationId?: string;
  /** Echoed back for caller's logs/audit. */
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  attachmentCount: number;
  /** Populated only when mode === 'dry-run' — the payload that WOULD POST. */
  payload?: SendMailPayload;
}

const MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024; // 30 MB practical /sendmail JSON cap

const MIME_BY_EXT: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.csv': 'text/csv',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.zip': 'application/zip',
  '.json': 'application/json',
  '.eml': 'message/rfc822',
};

export async function run(
  deps: SendMailDeps,
  opts: SendMailOptions = {},
): Promise<SendMailResult> {
  // -------- Validation --------
  const to = parseRecipients(opts.to);
  if (to.length === 0) {
    throw new UsageError('send-mail: --to is required (one or more recipients)');
  }
  const cc = parseRecipients(opts.cc);
  const bcc = parseRecipients(opts.bcc);

  if (typeof opts.subject !== 'string' || opts.subject.length === 0) {
    throw new UsageError('send-mail: --subject is required');
  }

  const hasHtml = typeof opts.html === 'string' && opts.html.length > 0;
  const hasText = typeof opts.text === 'string' && opts.text.length > 0;
  if (!hasHtml && !hasText) {
    throw new UsageError(
      'send-mail: at least one of --html <file> or --text <file> is required',
    );
  }

  // -------- Body load --------
  const reader = deps.readFile ?? ((p: string) => fs.readFile(p));
  let bodyContentType: SendBodyContentType;
  let bodyContent: string;
  if (hasHtml) {
    bodyContentType = 'HTML';
    bodyContent = await readBodyFile(reader, opts.html as string, '--html');
  } else {
    bodyContentType = 'Text';
    bodyContent = await readBodyFile(reader, opts.text as string, '--text');
  }

  // -------- Attachment load --------
  const attachmentPaths = Array.isArray(opts.attach) ? opts.attach : [];
  const attachments: SendFileAttachment[] = [];
  let totalAttachmentBytes = 0;
  for (const p of attachmentPaths) {
    const att = await loadAttachment(reader, p);
    totalAttachmentBytes += att.Size ?? 0;
    if (totalAttachmentBytes > MAX_ATTACHMENT_BYTES) {
      throw new UsageError(
        `send-mail: combined attachment size exceeds ${MAX_ATTACHMENT_BYTES} bytes ` +
          `(${totalAttachmentBytes} > limit). Split into multiple emails or use a ` +
          'shared link instead.',
      );
    }
    attachments.push(att);
  }

  // -------- Session + CC-self --------
  const session = await ensureSession(deps);
  const ccWithSelf = applyCcSelf(cc, session, opts.ccSelf);

  // -------- Build payload --------
  const payload: SendMailPayload = {
    Subject: opts.subject,
    Body: { ContentType: bodyContentType, Content: bodyContent },
    ToRecipients: to.map((addr) => ({ EmailAddress: { Address: addr } })),
  };
  if (ccWithSelf.length > 0) {
    payload.CcRecipients = ccWithSelf.map((addr) => ({
      EmailAddress: { Address: addr },
    }));
  }
  if (bcc.length > 0) {
    payload.BccRecipients = bcc.map((addr) => ({
      EmailAddress: { Address: addr },
    }));
  }
  if (attachments.length > 0) {
    payload.Attachments = attachments;
  }

  // -------- Dry-run short-circuit --------
  if (opts.dryRun === true) {
    return {
      mode: 'dry-run',
      to,
      cc: ccWithSelf,
      bcc,
      subject: opts.subject,
      attachmentCount: attachments.length,
      payload,
    };
  }

  // -------- Dispatch --------
  const client = deps.createClient(session);
  if (opts.sendNow === true) {
    try {
      await client.sendMail(payload, {
        saveToSentItems: opts.saveSent !== false,
      });
    } catch (err) {
      throw mapHttpError(err);
    }
    return {
      mode: 'sent',
      to,
      cc: ccWithSelf,
      bcc,
      subject: opts.subject,
      attachmentCount: attachments.length,
    };
  }

  // Default: create draft + activate Outlook desktop.
  let draft: CreateDraftResult;
  try {
    draft = await client.createDraft(payload);
  } catch (err) {
    throw mapHttpError(err);
  }

  if (opts.open !== false) {
    const activate = deps.activateOutlook ?? activateOutlookApp;
    try {
      await activate();
    } catch (err) {
      // Activation failure is non-fatal — the draft was created.
      const msg = err instanceof Error ? err.message : String(err);
      process.stderr.write(`send-mail: Outlook activation failed: ${msg}\n`);
    }
  }

  return {
    mode: 'draft',
    id: draft.Id,
    webLink: draft.WebLink,
    conversationId: draft.ConversationId,
    to,
    cc: ccWithSelf,
    bcc,
    subject: opts.subject,
    attachmentCount: attachments.length,
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Parse a recipient input that may be:
 *   - undefined  → []
 *   - string ("a@x.com, b@y.com")  → split on comma
 *   - string[]   → flatten + split each on comma (commander gives an array
 *                  when --to is repeated, but each value may itself contain
 *                  commas if the user mixed forms)
 *
 * Empty entries are skipped; whitespace is trimmed; minimum sanity check on
 * the address shape (must contain `@`).
 */
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
        throw new UsageError(
          `send-mail: invalid recipient address (no '@'): ${trimmed}`,
        );
      }
      out.push(trimmed);
    }
  }
  return out;
}

/**
 * Add the authenticated user's UPN to CC unless the caller opted out via
 * `ccSelf: false`. Avoids duplication if the user is already in TO/CC.
 *
 * NOTE: deduplication is case-insensitive on the local-part too (M365
 * addresses are case-insensitive).
 */
function applyCcSelf(
  cc: string[],
  session: SessionFile,
  ccSelfFlag: boolean | undefined,
): string[] {
  if (ccSelfFlag === false) return cc;
  const self = session.account?.upn;
  if (typeof self !== 'string' || self.length === 0) return cc;
  const lower = cc.map((a) => a.toLowerCase());
  if (lower.includes(self.toLowerCase())) return cc;
  return [...cc, self];
}

async function readBodyFile(
  reader: (p: string) => Promise<Buffer>,
  filePath: string,
  flagName: string,
): Promise<string> {
  try {
    const buf = await reader(filePath);
    return buf.toString('utf-8');
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new UsageError(
      `send-mail: ${flagName} file read failed (${filePath}): ${msg}`,
    );
  }
}

async function loadAttachment(
  reader: (p: string) => Promise<Buffer>,
  filePath: string,
): Promise<SendFileAttachment> {
  let buf: Buffer;
  try {
    buf = await reader(filePath);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new UsageError(
      `send-mail: --attach file read failed (${filePath}): ${msg}`,
    );
  }
  const name = path.basename(filePath);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_BY_EXT[ext] ?? 'application/octet-stream';
  return {
    '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
    Name: name,
    ContentType: contentType,
    ContentBytes: buf.toString('base64'),
    IsInline: false,
    Size: buf.length,
  };
}

