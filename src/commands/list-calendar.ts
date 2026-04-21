// src/commands/list-calendar.ts
//
// List upcoming calendar events.
// See project-design.md §2.13.6 and refined spec §5.6.

import type { CliConfig } from '../config/config';
import type { OutlookClient } from '../http/outlook-client';
import type { EventSummary, ODataListResponse } from '../http/types';
import type { SessionFile } from '../session/schema';

import { ensureSession, mapHttpError, UsageError } from './list-mail';

export interface ListCalendarDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (path: string) => Promise<SessionFile | null>;
  saveSession: (path: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  createClient: (s: SessionFile) => OutlookClient;
}

export interface ListCalendarOptions {
  from?: string;
  to?: string;
  tz?: string;
}

const MS_PER_DAY = 24 * 60 * 60 * 1000;

/**
 * Resolve a human keyword / ISO8601 string to a concrete ISO8601 timestamp.
 *
 *   "now"        → current instant (UTC)
 *   "now + Nd"   → current instant + N days (whitespace-insensitive)
 *   otherwise    → validated ISO8601 (throws UsageError on parse failure)
 */
export function resolveCalendarDate(raw: string, label: string): string {
  const trimmed = raw.trim();
  if (trimmed === 'now') {
    return new Date().toISOString();
  }
  const rel = trimmed.match(/^now\s*\+\s*(\d+)\s*d$/i);
  if (rel) {
    const days = Number.parseInt(rel[1]!, 10);
    return new Date(Date.now() + days * MS_PER_DAY).toISOString();
  }
  const t = Date.parse(trimmed);
  if (!Number.isFinite(t)) {
    throw new UsageError(
      `list-calendar: ${label} is not a valid ISO8601 date (got ${JSON.stringify(raw)})`,
    );
  }
  return new Date(t).toISOString();
}

export async function run(
  deps: ListCalendarDeps,
  opts: ListCalendarOptions = {},
): Promise<EventSummary[]> {
  const fromRaw = opts.from ?? deps.config.calFrom;
  const toRaw = opts.to ?? deps.config.calTo;

  const startDateTime = resolveCalendarDate(fromRaw, '--from');
  const endDateTime = resolveCalendarDate(toRaw, '--to');

  // `--tz` is accepted for forward-compatibility; the REST call itself carries
  // absolute UTC timestamps so the server's returned DateTimes are unaffected
  // by this flag. We merely validate when provided.
  const tz = opts.tz ?? deps.config.tz;
  if (typeof tz !== 'string' || tz.length === 0) {
    throw new UsageError('list-calendar: --tz must be a non-empty IANA zone');
  }

  const session = await ensureSession(deps);
  const client = deps.createClient(session);

  const query = {
    startDateTime,
    endDateTime,
    $orderby: 'Start/DateTime asc',
    $select: 'Id,Subject,Start,End,Organizer,Location,IsAllDay',
  };

  try {
    const resp = await client.get<ODataListResponse<EventSummary>>(
      '/api/v2.0/me/calendarview',
      query,
    );
    return Array.isArray(resp.value) ? resp.value : [];
  } catch (err) {
    throw mapHttpError(err);
  }
}
