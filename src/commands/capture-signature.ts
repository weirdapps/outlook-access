// src/commands/capture-signature.ts
//
// One-time capture of the user's email signature from a SentItems message.
// Heuristically extracts the signature block and saves to a known path so
// reply/forward commands can append it automatically.
//
// Heuristic priority:
//   1. <div id="Signature">...</div>  — Outlook web signature wrapper
//   2. <div class="elementToProof">...</div> — newer Outlook web variant
//   3. Last <hr>...remainder — fallback
//   4. Otherwise: split on common reply markers ("On <date>... wrote:",
//      "From: ...", "Sent: ...") and keep the first part minus visible
//      reply boilerplate.
//
// If none of the heuristics yields a non-empty result, returns the WHOLE
// message body — the user can hand-edit the file to trim down.

import { promises as fs } from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import type { CliConfig } from '../config/config';
import type { OutlookClient } from '../http/outlook-client';
import type { MessageSummary } from '../http/types';
import type { SessionFile } from '../session/schema';
import { extractCidReferences, saveSignatureAssets } from '../util/signature-assets';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface CaptureSignatureDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
  /** Test override for fs.writeFile. */
  writeFile?: (p: string, data: string) => Promise<void>;
  /** Test override for output dir resolution. */
  homeDir?: () => string;
}

export interface CaptureSignatureOptions {
  /** Override message id to extract from. Default: latest SentItems message. */
  fromMessage?: string;
  /** Override output path. Default: ~/.outlook-cli/signature.html */
  out?: string;
}

export interface CaptureSignatureResult {
  /** Where the signature was written. */
  path: string;
  /** Source message id used for extraction. */
  sourceMessageId: string;
  /** Source message subject (for context in the result output). */
  sourceSubject: string;
  /** The extracted HTML signature (also written to the file). */
  signature: string;
  /** Which heuristic produced the signature, for debugging. */
  heuristic: 'div-signature' | 'div-elementtoproof' | 'last-hr' | 'reply-marker' | 'whole-body';
  /** Where the signature inline image assets were saved (or null if none). */
  assetsDir: string | null;
  /** Number of inline images captured (matched the signature's cid: refs). */
  inlineAssetCount: number;
  /** cid: references in signature that had no matching attachment (orphans). */
  unmatchedCidRefs: string[];
}

const DEFAULT_OUT_REL = path.join('.outlook-cli', 'signature.html');

export async function run(
  deps: CaptureSignatureDeps,
  opts: CaptureSignatureOptions = {},
): Promise<CaptureSignatureResult> {
  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  // Pick the source message: explicit --from-message wins, else latest from SentItems.
  let sourceId: string;
  if (typeof opts.fromMessage === 'string' && opts.fromMessage.length > 0) {
    sourceId = opts.fromMessage;
  } else {
    let latest: MessageSummary[];
    try {
      latest = await client.listMessagesInFolder('SentItems', {
        top: 1,
        orderBy: 'SentDateTime desc',
        select: ['Id', 'Subject', 'SentDateTime'],
      });
    } catch (err) {
      throw mapHttpError(err);
    }
    if (latest.length === 0) {
      throw new UsageError(
        'capture-signature: SentItems folder is empty — cannot capture a signature. ' +
          'Send at least one mail (or pass --from-message <id>).',
      );
    }
    sourceId = latest[0]!.Id;
  }

  // Fetch the full body.
  let msg;
  try {
    msg = await client.getMessage(sourceId, {
      select: ['Id', 'Subject', 'Body'],
    });
  } catch (err) {
    throw mapHttpError(err);
  }
  const html = msg.Body?.Content ?? '';
  if (html.length === 0) {
    throw new UsageError(
      `capture-signature: source message ${sourceId} has empty Body — pass a different --from-message`,
    );
  }

  const { signature, heuristic } = extractSignature(html);

  // Resolve output path (~ expansion).
  const home = (deps.homeDir ?? os.homedir)();
  const outPath = opts.out ?? path.join(home, DEFAULT_OUT_REL);

  // Ensure parent dir exists.
  await fs.mkdir(path.dirname(outPath), { recursive: true, mode: 0o700 });

  const writer = deps.writeFile ?? ((p: string, d: string) => fs.writeFile(p, d, { mode: 0o600 }));
  await writer(outPath, signature);

  // -------- Inline image assets --------
  // Scan signature for cid: refs; if any, fetch the source message's
  // attachments, match by ContentId, save to <signature-dir>/signature-assets/.
  const cidRefs = extractCidReferences(signature);
  let assetsDir: string | null = null;
  let inlineAssetCount = 0;
  let unmatchedCidRefs: string[] = [];
  if (cidRefs.length > 0) {
    let attachments;
    try {
      attachments = await client.listMessageAttachments(sourceId);
    } catch (err) {
      throw mapHttpError(err);
    }
    // Filter to inline FileAttachments matching the cid refs.
    const wanted = new Set(cidRefs);
    const matchedAttachments = attachments.filter(
      (a) =>
        a['@odata.type'] === '#Microsoft.OutlookServices.FileAttachment' &&
        typeof a.ContentId === 'string' &&
        wanted.has(a.ContentId) &&
        typeof a.ContentBytes === 'string' &&
        a.ContentBytes.length > 0,
    );
    const matched = new Set(matchedAttachments.map((a) => a.ContentId as string));
    unmatchedCidRefs = cidRefs.filter((c) => !matched.has(c));

    if (matchedAttachments.length > 0) {
      assetsDir = path.join(path.dirname(outPath), 'signature-assets');
      await saveSignatureAssets({
        assetsDir,
        sourceMessageId: sourceId,
        attachments: matchedAttachments.map((a) => ({
          contentId: a.ContentId as string,
          contentType: a.ContentType ?? 'application/octet-stream',
          contentBytesBase64: a.ContentBytes as string,
          name: a.Name,
        })),
      });
      inlineAssetCount = matchedAttachments.length;
    }
  }

  return {
    path: outPath,
    sourceMessageId: sourceId,
    sourceSubject: msg.Subject,
    signature,
    heuristic,
    assetsDir,
    inlineAssetCount,
    unmatchedCidRefs,
  };
}

/**
 * Try heuristics in priority order. Returns the first non-empty signature
 * found. Last-resort fallback returns the whole body so the user has SOMETHING
 * to edit by hand.
 */
export function extractSignature(html: string): {
  signature: string;
  heuristic: CaptureSignatureResult['heuristic'];
} {
  // 1. <div id="Signature">...</div>
  const sigDiv = matchOuterDiv(html, /<div\s+id\s*=\s*["']?Signature["']?[^>]*>/i);
  if (sigDiv) return { signature: sigDiv.trim(), heuristic: 'div-signature' };

  // 2. <div class="elementToProof">...</div>
  const proofDiv = matchOuterDiv(
    html,
    /<div[^>]*class\s*=\s*["'][^"']*\belementToProof\b[^"']*["'][^>]*>/i,
  );
  if (proofDiv) return { signature: proofDiv.trim(), heuristic: 'div-elementtoproof' };

  // 3. Last <hr>... — split on the LAST <hr>, take everything after it.
  const hrIdx = html.toLowerCase().lastIndexOf('<hr');
  if (hrIdx >= 0) {
    // Find end of <hr ...> tag
    const tagEnd = html.indexOf('>', hrIdx);
    if (tagEnd > 0) {
      const after = html.slice(tagEnd + 1).trim();
      if (after.length > 0 && after.length < html.length / 2) {
        return { signature: after, heuristic: 'last-hr' };
      }
    }
  }

  // 4. Reply marker split — take content BEFORE the first reply marker.
  // "On <date> ... wrote:" / "From: " / "Sent: " / "-----Original Message-----"
  // Use lazy quantifiers to prevent ReDoS backtracking.
  const replyMarkerRe =
    /(<div[^>]*>\s*-{2,}\s*Original Message\s*-{2,}?)|(<p[^>]*>On\s+[^<]+?wrote:)|(<div[^>]*>From:\s)|(<p[^>]*>From:\s)/i;
  const m = html.match(replyMarkerRe);
  if (m && typeof m.index === 'number' && m.index > 0) {
    const before = html.slice(0, m.index).trim();
    if (before.length > 0) {
      // Keep the LAST paragraph block (signature usually sits in its own
      // <p>...</p> right above the reply marker). Find the position of the
      // last `<p` opening tag and slice from there.
      const lastPOpen = before.toLowerCase().lastIndexOf('<p');
      let sig: string;
      if (lastPOpen >= 0) {
        sig = before.slice(lastPOpen).trim();
      } else {
        // No <p> blocks — take everything before the marker.
        sig = before;
      }
      if (sig.length > 0) {
        return { signature: sig, heuristic: 'reply-marker' };
      }
    }
  }

  // 5. Last resort: return whole body.
  return { signature: html.trim(), heuristic: 'whole-body' };
}

/**
 * Find the FIRST div whose opening tag matches `openTagRe`, then return the
 * full element including its matched closing </div>. Handles nested divs by
 * counting depth.
 *
 * Returns `null` if no matching div is found or the match is unbalanced.
 */
function matchOuterDiv(html: string, openTagRe: RegExp): string | null {
  const m = html.match(openTagRe);
  if (!m || typeof m.index !== 'number') return null;
  const start = m.index;
  // Find end of opening tag
  const openEnd = html.indexOf('>', start);
  if (openEnd < 0) return null;

  let depth = 1;
  let i = openEnd + 1;
  const lower = html.toLowerCase();
  while (i < lower.length && depth > 0) {
    const nextOpen = lower.indexOf('<div', i);
    const nextClose = lower.indexOf('</div', i);
    if (nextClose < 0) return null;
    if (nextOpen >= 0 && nextOpen < nextClose) {
      depth++;
      i = nextOpen + 4;
    } else {
      depth--;
      i = nextClose + 6;
    }
  }
  if (depth !== 0) return null;
  return html.slice(start, i);
}

// Re-export for CLI + tests.
export { UsageError };
