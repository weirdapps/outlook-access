// src/commands/get-event.ts
//
// Retrieve a single calendar event.
// See project-design.md §2.13.7 and refined spec §5.7.

import type { CliConfig } from '../config/config';
import type { OutlookClient } from '../http/outlook-client';
import type { Event } from '../http/types';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';
import type { BodyMode } from './get-mail';

export interface GetEventDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface GetEventOptions {
  body?: BodyMode;
}

const BODY_MODES: readonly BodyMode[] = ['html', 'text', 'none'];

export async function run(
  deps: GetEventDeps,
  id: string,
  opts: GetEventOptions = {},
): Promise<Event> {
  if (typeof id !== 'string' || id.length === 0) {
    throw new UsageError('get-event: <id> is required');
  }

  const body: BodyMode = opts.body ?? deps.config.bodyMode;
  if (!BODY_MODES.includes(body)) {
    throw new UsageError(
      `get-event: --body must be one of ${BODY_MODES.join('|')} (got ${String(body)})`,
    );
  }

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const encodedId = encodeURIComponent(id);

  try {
    const event = await client.get<Event>(`/api/v2.0/me/events/${encodedId}`);
    // Body handling: HTML→text conversion deferred (matches get-mail). Only
    // "none" mutates the payload.
    if (body === 'none') {
      const copy: Event = { ...event };
      delete copy.Body;
      return copy;
    }
    return event;
  } catch (err) {
    throw mapHttpError(err);
  }
}
