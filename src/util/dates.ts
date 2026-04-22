// src/util/dates.ts
//
// Shared timestamp parser for the `--from` / `--to` window arguments accepted
// by list-calendar and list-mail. Grammar:
//
//   "now"              → current instant (UTC)
//   "now + Nd"         → current instant + N days (whitespace-insensitive)
//   "now - Nd"         → current instant − N days (whitespace-insensitive)
//   otherwise          → ISO8601 string, validated via Date.parse
//
// On failure the function does NOT throw directly — it returns a typed
// failure so each command can raise its own `UsageError` with the command-
// specific prefix the top-level handler expects.

const MS_PER_DAY = 24 * 60 * 60 * 1000;

export type TimestampParseResult =
  | { ok: true; iso: string }
  | { ok: false; reason: string };

/**
 * Parse a keyword / ISO8601 timestamp into an ISO8601 UTC string.
 *
 * Returns a discriminated result so the caller can format error messages
 * with the right command-level prefix.
 */
export function parseTimestamp(raw: string): TimestampParseResult {
  const trimmed = raw.trim();
  if (trimmed === 'now') {
    return { ok: true, iso: new Date().toISOString() };
  }
  const rel = trimmed.match(/^now\s*([+-])\s*(\d+)\s*d$/i);
  if (rel) {
    const sign = rel[1] === '-' ? -1 : 1;
    const days = Number.parseInt(rel[2]!, 10);
    const ts = Date.now() + sign * days * MS_PER_DAY;
    return { ok: true, iso: new Date(ts).toISOString() };
  }
  const t = Date.parse(trimmed);
  if (!Number.isFinite(t)) {
    return {
      ok: false,
      reason: `not a valid ISO8601 date (got ${JSON.stringify(raw)})`,
    };
  }
  return { ok: true, iso: new Date(t).toISOString() };
}
