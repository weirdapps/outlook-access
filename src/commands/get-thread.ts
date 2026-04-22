// src/commands/get-thread.ts
//
// Retrieve every message in a conversation (thread) regardless of folder.
// Input is a single message id (or `conv:<conversationId>` to skip the
// resolve hop); output is an array of messages ordered by ReceivedDateTime.
//
// See project-design.md §11 (Threads) and refined spec §5.5.

import type { CliConfig } from '../config/config';
import type { OutlookClient } from '../http/outlook-client';
import type { Message, MessageSummary } from '../http/types';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface GetThreadDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export type ThreadBodyMode = 'html' | 'text' | 'none';
export type ThreadOrder = 'asc' | 'desc';

export interface GetThreadOptions {
  body?: ThreadBodyMode;
  order?: ThreadOrder;
}

const BODY_MODES: readonly ThreadBodyMode[] = ['html', 'text', 'none'];
const ORDER_VALUES: readonly ThreadOrder[] = ['asc', 'desc'];

/**
 * Field set requested for each message in the thread. Body is added
 * conditionally (when --body != "none").
 */
const THREAD_BASE_SELECT = [
  'Id',
  'Subject',
  'From',
  'ReceivedDateTime',
  'HasAttachments',
  'IsRead',
  'WebLink',
  'ConversationId',
  'ParentFolderId',
  'SentDateTime',
];

/** The `conv:` prefix that lets the caller pass a ConversationId directly. */
const CONV_PREFIX = 'conv:';

export interface ThreadResult {
  conversationId: string;
  /** Number of messages in the thread. Equivalent to `messages.length`. */
  count: number;
  messages: MessageSummary[];
}

export async function run(
  deps: GetThreadDeps,
  idOrConv: string,
  opts: GetThreadOptions = {},
): Promise<ThreadResult> {
  if (typeof idOrConv !== 'string' || idOrConv.length === 0) {
    throw new UsageError('get-thread: <id> is required');
  }

  const body: ThreadBodyMode = opts.body ?? 'text';
  if (!BODY_MODES.includes(body)) {
    throw new UsageError(
      `get-thread: --body must be one of ${BODY_MODES.join('|')} (got ${String(body)})`,
    );
  }

  const order: ThreadOrder = opts.order ?? 'asc';
  if (!ORDER_VALUES.includes(order)) {
    throw new UsageError(
      `get-thread: --order must be one of ${ORDER_VALUES.join('|')} (got ${String(order)})`,
    );
  }

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  // Step 1: resolve the conversation id. Two input modes:
  //   - `conv:<raw>` → use the suffix directly.
  //   - anything else → assume it's a message id; fetch ConversationId
  //     with a tight $select.
  let conversationId: string;
  try {
    if (idOrConv.startsWith(CONV_PREFIX)) {
      const raw = idOrConv.slice(CONV_PREFIX.length);
      if (raw.length === 0) {
        throw new UsageError(
          'get-thread: conv: prefix requires a non-empty conversation id',
        );
      }
      conversationId = raw;
    } else {
      const msg = await client.get<Message>(
        `/api/v2.0/me/messages/${encodeURIComponent(idOrConv)}`,
        { $select: 'Id,ConversationId' },
      );
      if (typeof msg.ConversationId !== 'string' || msg.ConversationId.length === 0) {
        throw new UsageError(
          `get-thread: message ${idOrConv} has no ConversationId (upstream returned an empty field)`,
        );
      }
      conversationId = msg.ConversationId;
    }
  } catch (err) {
    if (err instanceof UsageError) {
      throw err;
    }
    throw mapHttpError(err);
  }

  // Step 2: fetch every message in the conversation.
  const select = [...THREAD_BASE_SELECT];
  if (body !== 'none') {
    select.push('Body');
    select.push('BodyPreview');
  }

  try {
    const messages = await client.listMessagesByConversation(conversationId, {
      select,
      orderBy: `ReceivedDateTime ${order}`,
    });
    return {
      conversationId,
      count: messages.length,
      messages,
    };
  } catch (err) {
    throw mapHttpError(err);
  }
}
