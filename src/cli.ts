#!/usr/bin/env -S node --preserve-symlinks
// src/cli.ts
//
// Commander bootstrap for the Outlook CLI.
// See project-design.md §2.14 and refined-request-outlook-cli.md §5.

import { Command, Option } from 'commander';

import { loadConfig, type CliConfig, type CliFlags, type BodyMode } from './config/config';
import {
  AuthError as CliAuthError,
  ConfigurationError,
  IoError,
  OutlookCliError,
  UpstreamError,
} from './config/errors';
import {
  AuthCaptureError,
  captureOutlookSession,
  type CaptureResult,
} from './auth/browser-capture';
import { createOutlookClient, type OutlookClient } from './http/outlook-client';
import { CollisionError } from './http/errors';
import { loadSession, saveSession } from './session/store';
import type { SessionFile } from './session/schema';
import { formatOutput, type ColumnSpec, type OutputMode } from './output/formatter';
import type { EventSummary, MessageSummary } from './http/types';
import type {
  CreateFolderResult,
  CreateFolderSegment,
  MoveEntry,
  MoveFailedEntry,
  MoveMailResult,
} from './folders/types';

import * as authCheck from './commands/auth-check';
import * as authRenew from './commands/auth-renew';
import * as login from './commands/login';
import * as listMail from './commands/list-mail';
import * as getMail from './commands/get-mail';
import * as getThread from './commands/get-thread';
import * as downloadAttachments from './commands/download-attachments';
import * as downloadSharepointLink from './commands/download-sharepoint-link';
import * as listCalendar from './commands/list-calendar';
import * as getEvent from './commands/get-event';
import * as listFolders from './commands/list-folders';
import { LIST_FOLDERS_COLUMNS } from './commands/list-folders';
import * as findFolder from './commands/find-folder';
import * as createFolder from './commands/create-folder';
import * as moveMail from './commands/move-mail';
import * as sendMail from './commands/send-mail';
import * as captureSignature from './commands/capture-signature';
import * as reply from './commands/reply';

// ---------------------------------------------------------------------------
// Package version (read lazily so --help doesn't need it)
// ---------------------------------------------------------------------------

function readPackageVersion(): string {
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const pkg = require('../package.json') as { version?: string };
    return typeof pkg.version === 'string' ? pkg.version : '0.0.0';
  } catch {
    return '0.0.0';
  }
}

// ---------------------------------------------------------------------------
// Dependency wiring
// ---------------------------------------------------------------------------

interface CommandDeps {
  config: CliConfig;
  sessionPath: string;
  loadSession: (p: string) => Promise<SessionFile | null>;
  saveSession: (p: string, s: SessionFile) => Promise<void>;
  doAuthCapture: () => Promise<SessionFile>;
  /** Like doAuthCapture but also captures + persists a SharePoint session
   *  from the same Playwright context. Used by `login --sharepoint-host`. */
  doAuthCaptureWithSharepoint: (host: string) => Promise<{
    session: SessionFile;
    sharepointPath: string;
  }>;
  createClient: (s: SessionFile) => OutlookClient;
}

/** Convert a raw CaptureResult from Playwright into a persisted SessionFile. */
function captureToSessionFile(r: CaptureResult): SessionFile {
  return {
    version: 1,
    capturedAt: new Date().toISOString(),
    account: r.account,
    bearer: r.bearer,
    cookies: r.cookies,
    anchorMailbox: r.anchorMailbox,
  };
}

function buildDeps(globalFlags: CliFlags): CommandDeps {
  const config = loadConfig(globalFlags);
  const sessionPath = config.sessionFileOverride ?? config.sessionFilePath;

  // Re-auth closure. The client uses it on 401 unless `noAutoReauth` is set.
  const doAuthCapture = async (): Promise<SessionFile> => {
    const captured = await captureOutlookSession({
      profileDir: config.profileDir,
      chromeChannel: config.chromeChannel,
      loginTimeoutMs: config.loginTimeoutMs,
    });
    const sessionFile = captureToSessionFile(captured);
    await saveSession(sessionPath, sessionFile);
    return sessionFile;
  };

  // SharePoint-aware variant. Captures both Outlook + SharePoint from the
  // same persistent context, then persists each file independently.
  const doAuthCaptureWithSharepoint = async (
    host: string,
  ): Promise<{ session: SessionFile; sharepointPath: string }> => {
    const captured = await captureOutlookSession({
      profileDir: config.profileDir,
      chromeChannel: config.chromeChannel,
      loginTimeoutMs: config.loginTimeoutMs,
      sharepointHost: host,
    });
    const sessionFile = captureToSessionFile(captured);
    await saveSession(sessionPath, sessionFile);
    if (!captured.sharepoint) {
      throw new Error(`captureOutlookSession returned no sharepoint session for host "${host}"`);
    }
    const { defaultSharepointSessionPath, saveSharepointSession } =
      await import('./session/sharepoint-schema');
    const sharepointPath = defaultSharepointSessionPath();
    await saveSharepointSession(sharepointPath, captured.sharepoint);
    return { session: sessionFile, sharepointPath };
  };

  const createClient = (s: SessionFile): OutlookClient =>
    createOutlookClient({
      session: s,
      httpTimeoutMs: config.httpTimeoutMs,
      noAutoReauth: config.noAutoReauth,
      onReauthNeeded: doAuthCapture,
    });

  return {
    config,
    sessionPath,
    loadSession,
    saveSession,
    doAuthCapture,
    doAuthCaptureWithSharepoint,
    createClient,
  };
}

// ---------------------------------------------------------------------------
// Global-flag → CliFlags mapping
// ---------------------------------------------------------------------------

interface GlobalOpts {
  timeout?: string;
  loginTimeout?: string;
  chromeChannel?: string;
  sessionFile?: string;
  profileDir?: string;
  tz?: string;
  json?: boolean;
  table?: boolean;
  quiet?: boolean;
  autoReauth?: boolean; // commander's `--no-auto-reauth` gives `autoReauth:false`
  logFile?: string;
}

function parseIntOrUndef(v: string | undefined, label: string): number | undefined {
  if (v === undefined) return undefined;
  if (!/^-?\d+$/.test(v)) {
    throw new CommanderLikeError(`${label} must be an integer (got ${JSON.stringify(v)})`);
  }
  return Number.parseInt(v, 10);
}

function resolveOutputMode(g: GlobalOpts): OutputMode {
  // `--json` has `.default(true)`, so it is always set. Mutual-exclusion only
  // fires when the user EXPLICITLY passed both flags (via the raw argv).
  const argv = process.argv.slice(2);
  const explicitJson = argv.includes('--json');
  const explicitTable = argv.includes('--table') || g.table === true;
  if (explicitJson && explicitTable) {
    throw new CommanderLikeError('--json and --table are mutually exclusive');
  }
  return explicitTable ? 'table' : 'json';
}

function globalOptsToFlags(g: GlobalOpts): CliFlags {
  const flags: CliFlags = {};
  const timeout = parseIntOrUndef(g.timeout, '--timeout');
  if (timeout !== undefined) flags.httpTimeoutMs = timeout;
  const loginTimeout = parseIntOrUndef(g.loginTimeout, '--login-timeout');
  if (loginTimeout !== undefined) flags.loginTimeoutMs = loginTimeout;
  if (typeof g.chromeChannel === 'string') flags.chromeChannel = g.chromeChannel;
  if (typeof g.sessionFile === 'string') flags.sessionFileOverride = g.sessionFile;
  if (typeof g.profileDir === 'string') flags.profileDir = g.profileDir;
  if (typeof g.tz === 'string') flags.tz = g.tz;
  flags.outputMode = g.table === true ? 'table' : 'json';
  if (g.quiet === true) flags.quiet = true;
  // commander's `--no-auto-reauth` flips `autoReauth` to false.
  if (g.autoReauth === false) flags.noAutoReauth = true;
  if (typeof g.logFile === 'string') flags.logFilePath = g.logFile;
  return flags;
}

// ---------------------------------------------------------------------------
// Local usage error (mapped to exit code 2)
// ---------------------------------------------------------------------------

class CommanderLikeError extends Error {
  public readonly exitCode: number = 2;
  public readonly code: string = 'USAGE_ERROR';
  constructor(message: string) {
    super(message);
    this.name = 'CommanderLikeError';
  }
}

// ---------------------------------------------------------------------------
// Column definitions for --table mode
// ---------------------------------------------------------------------------

const LIST_MAIL_COLUMNS: ColumnSpec<MessageSummary>[] = [
  {
    header: 'Received',
    extract: (r) => r.ReceivedDateTime ?? '',
    maxWidth: 20,
  },
  {
    header: 'From',
    extract: (r) => r.From?.EmailAddress?.Name || r.From?.EmailAddress?.Address || '',
    maxWidth: 28,
  },
  {
    header: 'Subject',
    extract: (r) => r.Subject ?? '',
    maxWidth: 48,
  },
  {
    header: 'Att',
    extract: (r) => (r.HasAttachments ? 'yes' : 'no'),
  },
  {
    header: 'Id',
    extract: (r) => r.Id ?? '',
    // No maxWidth: IDs must never be truncated, otherwise copy-paste into
    // `get-mail`/`download-attachments` sends an ellipsis character instead
    // of the real bytes and the server returns ErrorInvalidIdMalformed.
  },
];

const GET_THREAD_COLUMNS: ColumnSpec<MessageSummary>[] = [
  {
    header: 'Received',
    extract: (r) => r.ReceivedDateTime ?? '',
    maxWidth: 20,
  },
  {
    header: 'From',
    extract: (r) => r.From?.EmailAddress?.Name || r.From?.EmailAddress?.Address || '',
    maxWidth: 28,
  },
  {
    header: 'Subject',
    extract: (r) => r.Subject ?? '',
    maxWidth: 56,
  },
  {
    header: 'Id',
    extract: (r) => r.Id ?? '',
  },
];

const LIST_CALENDAR_COLUMNS: ColumnSpec<EventSummary>[] = [
  {
    header: 'Start',
    extract: (r) => r.Start?.DateTime ?? '',
    maxWidth: 20,
  },
  {
    header: 'End',
    extract: (r) => r.End?.DateTime ?? '',
    maxWidth: 20,
  },
  {
    header: 'Subject',
    extract: (r) => r.Subject ?? '',
    maxWidth: 40,
  },
  {
    header: 'Organizer',
    extract: (r) => r.Organizer?.EmailAddress?.Name || r.Organizer?.EmailAddress?.Address || '',
    maxWidth: 28,
  },
  {
    header: 'Location',
    extract: (r) => r.Location?.DisplayName ?? '',
    maxWidth: 28,
  },
  {
    header: 'Id',
    extract: (r) => r.Id ?? '',
    // No maxWidth: event IDs must stay intact for copy-paste into `get-event`.
  },
];

// `create-folder` table columns are applied to `result.created[]` (see
// project-design §10.7). The shape is `CreateFolderSegment`.
const CREATE_FOLDER_COLUMNS: ColumnSpec<CreateFolderSegment>[] = [
  {
    header: 'Path',
    extract: (r) => r.Path ?? '',
    maxWidth: 48,
  },
  {
    header: 'Id',
    extract: (r) => r.Id ?? '',
    // No maxWidth: folder IDs must stay intact for copy-paste.
  },
  {
    header: 'PreExisting',
    extract: (r) => (r.PreExisting ? 'yes' : 'no'),
  },
];

/**
 * `move-mail` table rows — union of `MoveEntry` (success) and
 * `MoveFailedEntry` (failure under `--continue-on-error`). Columns per
 * §10.7: `Source Id | New Id | Status | Error`.
 */
interface MoveMailRow {
  sourceId: string;
  newId?: string;
  status: 'moved' | 'failed';
  error?: string;
}

const MOVE_MAIL_COLUMNS: ColumnSpec<MoveMailRow>[] = [
  {
    header: 'Source Id',
    extract: (r) => r.sourceId ?? '',
    // No maxWidth: message ids must stay intact.
  },
  {
    header: 'New Id',
    extract: (r) => r.newId ?? '',
    // No maxWidth: message ids must stay intact.
  },
  {
    header: 'Status',
    extract: (r) => r.status ?? '',
  },
  {
    header: 'Error',
    extract: (r) => r.error ?? '',
    maxWidth: 48,
  },
];

/** Flatten a `MoveMailResult` into the `MoveMailRow[]` shape expected by
 *  `MOVE_MAIL_COLUMNS`. Successes come first, failures after. */
function toMoveMailRows(r: MoveMailResult): MoveMailRow[] {
  const rows: MoveMailRow[] = [];
  for (const m of r.moved as MoveEntry[]) {
    rows.push({ sourceId: m.sourceId, newId: m.newId, status: 'moved' });
  }
  for (const f of r.failed as MoveFailedEntry[]) {
    const parts: string[] = [];
    if (f.error?.code) parts.push(f.error.code);
    if (typeof f.error?.httpStatus === 'number') parts.push(`HTTP ${f.error.httpStatus}`);
    if (f.error?.message) parts.push(f.error.message);
    rows.push({
      sourceId: f.sourceId,
      status: 'failed',
      error: parts.join(' — '),
    });
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Output helpers
// ---------------------------------------------------------------------------

function emitResult(data: unknown, mode: OutputMode, columns?: ColumnSpec<unknown>[]): void {
  // Table is only meaningful for array-ish results that match the column
  // spec. For non-tabular results we fall back to JSON silently.
  if (mode === 'table' && columns) {
    process.stdout.write(formatOutput(data, 'table', columns) + '\n');
    return;
  }
  process.stdout.write(JSON.stringify(data, null, 2) + '\n');
}

// ---------------------------------------------------------------------------
// Error handling
// ---------------------------------------------------------------------------

interface ErrorPayload {
  error: {
    code: string;
    message: string;
    [key: string]: unknown;
  };
}

function formatErrorJson(err: unknown): ErrorPayload {
  // NOSONAR S3776 - error classification logic
  if (err instanceof ConfigurationError) {
    return {
      error: {
        code: err.code,
        message: err.message,
        missingSetting: err.missingSetting,
        checkedSources: err.checkedSources,
      },
    };
  }
  if (err instanceof CliAuthError) {
    return { error: { code: err.code, message: err.message } };
  }
  if (err instanceof UpstreamError) {
    const payload: ErrorPayload = {
      error: { code: err.code, message: err.message },
    };
    if (err.httpStatus !== undefined) payload.error.httpStatus = err.httpStatus;
    if (err.requestId !== undefined) payload.error.requestId = err.requestId;
    if (err.url !== undefined) payload.error.url = err.url;
    return payload;
  }
  if (err instanceof IoError) {
    const payload: ErrorPayload = {
      error: { code: err.code, message: err.message },
    };
    if (err.path !== undefined) payload.error.path = err.path;
    return payload;
  }
  if (err instanceof CollisionError) {
    const payload: ErrorPayload = {
      error: { code: err.code, message: err.message },
    };
    if (err.path !== undefined) payload.error.path = err.path;
    if (err.parentId !== undefined) payload.error.parentId = err.parentId;
    return payload;
  }
  if (err instanceof AuthCaptureError) {
    return { error: { code: err.code, message: err.message } };
  }
  if (err instanceof OutlookCliError) {
    return { error: { code: err.code, message: err.message } };
  }
  if (err instanceof CommanderLikeError) {
    return { error: { code: err.code, message: err.message } };
  }
  // Commander usage errors — surface the commander code under a USAGE_ERROR label.
  if (err && typeof err === 'object' && 'code' in err) {
    const code = (err as { code: unknown }).code;
    if (typeof code === 'string' && code.startsWith('commander.')) {
      return {
        error: {
          code: 'USAGE_ERROR',
          message: String((err as { message?: unknown }).message ?? err),
          commanderCode: code,
        },
      };
    }
  }
  return {
    error: {
      code: 'UNEXPECTED',
      message: String((err as Error)?.message ?? err),
    },
  };
}

function exitCodeFor(err: unknown): number {
  if (err instanceof ConfigurationError) return 3;
  if (err instanceof CliAuthError) return 4;
  if (err instanceof AuthCaptureError) return 4;
  if (err instanceof UpstreamError) return 5;
  if (err instanceof IoError) return 6;
  if (err instanceof CollisionError) return 6;
  if (err instanceof OutlookCliError) return err.exitCode;
  if (err instanceof CommanderLikeError) return err.exitCode;
  // Commander usage errors (unknown command/option, missing arg, etc.) must
  // map to exit 2 per the error taxonomy in design §4.
  if (err && typeof err === 'object' && 'code' in err) {
    const code = (err as { code: unknown }).code;
    if (typeof code === 'string' && code.startsWith('commander.')) {
      return 2;
    }
  }
  // Other objects with a numeric exitCode (e.g., process-level errors).
  if (
    err &&
    typeof err === 'object' &&
    'exitCode' in err &&
    typeof (err as { exitCode: unknown }).exitCode === 'number'
  ) {
    return (err as { exitCode: number }).exitCode;
  }
  return 1;
}

function reportError(err: unknown): number {
  const exit = exitCodeFor(err);
  const payload = formatErrorJson(err);
  try {
    process.stderr.write(JSON.stringify(payload, null, 2) + '\n');
  } catch {
    // Never fail the exit path on stderr I/O problems.
  }
  return exit;
}

// ---------------------------------------------------------------------------
// Commander program
// ---------------------------------------------------------------------------

type ActionHandler<O, Args extends unknown[]> = (
  deps: CommandDeps,
  globalOpts: GlobalOpts,
  cmdOpts: O,
  ...args: Args
) => Promise<void>;

/**
 * Thin wrapper so that every command action:
 *   - builds deps (loads config)
 *   - catches thrown errors
 *   - sets process.exitCode
 */
function makeAction<O, Args extends unknown[]>(
  program: Command,
  handler: ActionHandler<O, Args>,
): (...args: [...Args, O, Command]) => Promise<void> {
  return async (...args: [...Args, O, Command]): Promise<void> => {
    const cmdOpts = args[args.length - 2] as O;
    const positional = args.slice(0, args.length - 2) as Args;
    const globalOpts = program.opts() as GlobalOpts;
    try {
      const flags = globalOptsToFlags(globalOpts);
      const deps = buildDeps(flags);
      await handler(deps, globalOpts, cmdOpts, ...positional);
    } catch (err) {
      process.exitCode = reportError(err);
    }
  };
}

export async function main(argv: string[]): Promise<number> {
  const program = new Command();
  program
    .name('outlook-cli')
    .description('CLI tool for reading Outlook mail and calendar via outlook.office.com/api/v2.0')
    .version(readPackageVersion())
    // Global flags. No `enablePositionalOptions()` so users can place global
    // options either before or after the subcommand (e.g.
    // `list-mail -n 5 --table`). Subcommand options never clash by name.

    .option('--timeout <ms>', 'Per-REST-call HTTP timeout (default 30000)')
    .option('--login-timeout <ms>', 'Max wait for interactive login (default 300000)')
    .option('--chrome-channel <name>', 'Playwright Chrome channel (default "chrome")')
    .option('--session-file <path>', 'Override session file path')
    .option('--profile-dir <path>', 'Override Playwright profile directory')
    .option('--tz <iana>', 'IANA timezone override')
    .addOption(new Option('--json', 'Emit JSON to stdout (default)').default(true))
    .option('--table', 'Emit a human-readable table to stdout')
    .option('--quiet', 'Suppress stderr progress messages', false)
    .option('--no-auto-reauth', 'Do not auto-reopen the browser on 401 or expired session')
    .option('--log-file <path>', 'Write debug log to a file (mode 0600)');

  // -------- login --------
  program
    .command('login')
    .description('Open Chrome and capture a fresh Outlook session')
    .option('--force', 'Ignore any cached session and always open the browser', false)
    .option(
      '--sharepoint-host <host>',
      'After Outlook login, also capture a SharePoint session for this host (e.g. nbg.sharepoint.com)',
    )
    .action(
      makeAction<{ force?: boolean; sharepointHost?: string }, []>(
        program,
        async (deps, g, cmdOpts) => {
          const result = await login.run(deps, {
            force: cmdOpts.force === true,
            sharepointHost: cmdOpts.sharepointHost,
          });
          emitResult(result, resolveOutputMode(g));
        },
      ),
    );

  // -------- auth-check --------
  program
    .command('auth-check')
    .description('Verify the cached session is present and accepted by Outlook')
    .action(
      makeAction<Record<string, never>, []>(program, async (deps, g) => {
        const result = await authCheck.run(deps);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- auth-renew --------
  program
    .command('auth-renew')
    .description('Silently renew the bearer using the persisted browser profile (headless)')
    .option('--timeout <ms>', 'Headless capture timeout (default 30000)', parseIntArg)
    .action(
      makeAction<{ timeout?: number }, []>(program, async (deps, g, cmdOpts) => {
        const result = await authRenew.run(deps, {
          timeoutMs: cmdOpts.timeout,
        });
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- list-mail --------
  program
    .command('list-mail')
    .description('List recent messages from a well-known folder')
    .option('-n, --top <N>', 'Number of messages (1..1000, default 10)', parseIntArg)
    .option(
      '--folder <name>',
      'Folder name (Inbox|SentItems|Drafts|DeletedItems|Archive or path/alias)',
    )
    .option('--folder-id <id>', 'Raw folder id (XOR with --folder)')
    .option(
      '--folder-parent <spec>',
      'Anchor for a path/bare-name in --folder (default MsgFolderRoot)',
    )
    .option('--select <csv>', 'Comma-separated $select fields')
    .option('--since <iso>', 'ISO-8601 UTC: include only messages with ReceivedDateTime >= this')
    .option('--until <iso>', 'ISO-8601 UTC: include only messages with ReceivedDateTime < this')
    .option(
      '--from <iso|keyword>',
      'Lower bound (ge) on ReceivedDateTime. ISO-8601 or now / now+Nd / now-Nd. Mutually exclusive with --since.',
    )
    .option(
      '--to <iso|keyword>',
      'Upper bound (lt) on ReceivedDateTime. Same grammar as --from. Mutually exclusive with --until.',
    )
    .option('--all', 'Auto-paginate via @odata.nextLink until exhausted', false)
    .option('--max <N>', 'Safety cap for --all (default 10000, max 100000)', parseIntArg)
    .option(
      '--just-count',
      'Return only {count, exact} via server-side $count=true. Ignores --top/--select. Mutually exclusive with --all.',
      false,
    )
    .action(
      makeAction<
        {
          top?: number;
          folder?: string;
          folderId?: string;
          folderParent?: string;
          select?: string;
          since?: string;
          until?: string;
          from?: string;
          to?: string;
          all?: boolean;
          max?: number;
          justCount?: boolean;
        },
        []
      >(program, async (deps, g, cmdOpts) => {
        const result = await listMail.run(deps, cmdOpts);
        const mode = resolveOutputMode(g);
        // --just-count returns {count, exact}, not a message array — emit as
        // a plain object regardless of mode (no columns).
        if (cmdOpts.justCount === true) {
          emitResult(result, mode);
        } else {
          emitResult(result, mode, LIST_MAIL_COLUMNS as unknown as ColumnSpec<unknown>[]);
        }
      }),
    );

  // -------- get-mail <id> --------
  program
    .command('get-mail')
    .argument('<id>', 'Message id')
    .description('Retrieve one message with optional body')
    .option('--body <mode>', 'Body inclusion: html|text|none')
    .action(
      makeAction<{ body?: BodyMode }, [string]>(program, async (deps, g, cmdOpts, id) => {
        const result = await getMail.run(deps, id, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- get-thread <id> --------
  program
    .command('get-thread')
    .argument('<id>', 'Message id (or "conv:<conversationId>" to skip the resolve hop)')
    .description('Retrieve every message in a conversation (thread) regardless of folder')
    .option('--body <mode>', 'Body inclusion: html|text|none (default text)')
    .option('--order <asc|desc>', 'ReceivedDateTime order (default asc = oldest first)')
    .action(
      makeAction<{ body?: getThread.ThreadBodyMode; order?: getThread.ThreadOrder }, [string]>(
        program,
        async (deps, g, cmdOpts, id) => {
          const result = await getThread.run(deps, id, cmdOpts);
          const mode = resolveOutputMode(g);
          if (mode === 'table') {
            emitResult(
              result.messages,
              mode,
              GET_THREAD_COLUMNS as unknown as ColumnSpec<unknown>[],
            );
          } else {
            emitResult(result, mode);
          }
        },
      ),
    );

  // -------- download-attachments <id> --------
  program
    .command('download-attachments')
    .argument('<id>', 'Message id')
    .description('Save all non-inline attachments from a message to a directory')
    .option('--out <dir>', 'Output directory (no default — must be provided)')
    .option('--overwrite', 'Overwrite existing files', false)
    .option('--include-inline', 'Include inline attachments', false)
    .action(
      makeAction<{ out?: string; overwrite?: boolean; includeInline?: boolean }, [string]>(
        program,
        async (deps, g, cmdOpts, id) => {
          const result = await downloadAttachments.run(deps, id, cmdOpts);
          emitResult(result, resolveOutputMode(g));
        },
      ),
    );

  // -------- download-sharepoint-link <url> --------
  program
    .command('download-sharepoint-link')
    .argument('<url>', 'SharePoint URL (typically from a ReferenceAttachment.SourceUrl)')
    .description('Fetch a SharePoint URL using the captured SharePoint session')
    .option('--out <dir>', 'Output directory (no default — must be provided)')
    .option('--overwrite', 'Overwrite existing files', false)
    .action(
      makeAction<{ out?: string; overwrite?: boolean }, [string]>(
        program,
        async (deps, g, cmdOpts, url) => {
          const result = await downloadSharepointLink.run(
            { httpTimeoutMs: deps.config.httpTimeoutMs },
            url,
            cmdOpts,
          );
          emitResult(result, resolveOutputMode(g));
        },
      ),
    );

  // -------- list-calendar --------
  program
    .command('list-calendar')
    .description('List upcoming calendar events within a window')
    .option('--from <iso>', 'Window start (ISO8601). Default: now')
    .option('--to <iso>', 'Window end (ISO8601). Default: now + 7d')
    .option('--tz <iana>', 'Timezone override')
    .action(
      makeAction<{ from?: string; to?: string; tz?: string }, []>(
        program,
        async (deps, g, cmdOpts) => {
          const result = await listCalendar.run(deps, cmdOpts);
          emitResult(
            result,
            resolveOutputMode(g),
            LIST_CALENDAR_COLUMNS as unknown as ColumnSpec<unknown>[],
          );
        },
      ),
    );

  // -------- get-event <id> --------
  program
    .command('get-event')
    .argument('<id>', 'Event id')
    .description('Retrieve one calendar event with optional body')
    .option('--body <mode>', 'Body inclusion: html|text|none')
    .action(
      makeAction<{ body?: BodyMode }, [string]>(program, async (deps, g, cmdOpts, id) => {
        const result = await getEvent.run(deps, id, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- list-folders --------
  program
    .command('list-folders')
    .description('List mail folders under a parent (well-known, path, or id:...)')
    .option(
      '--parent <spec>',
      'Parent folder (well-known alias, path, or id:<raw>). Default: MsgFolderRoot',
    )
    .option('--top <N>', 'Page size (1..250, default 100)', parseIntArg)
    .option('--recursive', 'Walk the full sub-tree (bounded)', false)
    .option('--include-hidden', 'Include folders whose IsHidden === true', false)
    .option(
      '--first-match',
      'On ambiguity, pick the oldest candidate (CreatedDateTime asc, Id asc)',
      false,
    )
    .action(
      makeAction<
        {
          parent?: string;
          top?: number;
          recursive?: boolean;
          includeHidden?: boolean;
          firstMatch?: boolean;
        },
        []
      >(program, async (deps, g, cmdOpts) => {
        const result = await listFolders.run(deps, cmdOpts);
        emitResult(
          result,
          resolveOutputMode(g),
          LIST_FOLDERS_COLUMNS as unknown as ColumnSpec<unknown>[],
        );
      }),
    );

  // -------- find-folder <spec> --------
  program
    .command('find-folder')
    .argument('<spec>', 'Folder query: well-known alias, path, or id:<raw>')
    .description('Resolve a folder query to a single ResolvedFolder')
    .option(
      '--anchor <spec>',
      'Anchor for path-form queries (well-known, path, or id:<raw>). Default: MsgFolderRoot',
    )
    .option(
      '--first-match',
      'On ambiguity, pick the oldest candidate (CreatedDateTime asc, Id asc)',
      false,
    )
    .action(
      makeAction<{ anchor?: string; firstMatch?: boolean }, [string]>(
        program,
        async (deps, g, cmdOpts, spec) => {
          const result = await findFolder.run(deps, spec, cmdOpts);
          // find-folder returns a single object — no ColumnSpec; emitResult
          // falls back to JSON per project-design §10.7.
          emitResult(result, resolveOutputMode(g));
        },
      ),
    );

  // -------- create-folder <path> --------
  program
    .command('create-folder')
    .argument('<path-or-name>', 'Folder name or slash-separated path to create')
    .description('Create (or idempotently reuse) a mail folder under an anchor')
    .option(
      '--parent <spec>',
      'Anchor folder (well-known, path, or id:<raw>). Default: MsgFolderRoot',
    )
    .option('--create-parents', 'Create missing intermediate segments on a nested path', false)
    .option('--idempotent', 'Return the existing folder on collision instead of exit 6', false)
    .action(
      makeAction<
        {
          parent?: string;
          createParents?: boolean;
          idempotent?: boolean;
        },
        [string]
      >(program, async (deps, g, cmdOpts, path) => {
        const result = await createFolder.run(deps, path, cmdOpts);
        // Table mode renders `result.created[]` with CREATE_FOLDER_COLUMNS;
        // JSON mode emits the full CreateFolderResult.
        const mode = resolveOutputMode(g);
        if (mode === 'table') {
          emitResult(
            (result as CreateFolderResult).created,
            mode,
            CREATE_FOLDER_COLUMNS as unknown as ColumnSpec<unknown>[],
          );
        } else {
          emitResult(result, mode);
        }
      }),
    );

  // -------- move-mail <messageIds...> --------
  program
    .command('move-mail')
    .argument('<messageIds...>', 'One or more source message ids to move')
    .description('Move one or more messages to a destination folder (returns new ids per §10.8)')
    .requiredOption('--to <spec>', 'Destination folder (well-known, path, or id:<raw>)')
    .option(
      '--first-match',
      'On ambiguity in --to, pick the oldest candidate (CreatedDateTime asc, Id asc)',
      false,
    )
    .option(
      '--continue-on-error',
      'Collect per-message failures into failed[] instead of aborting',
      false,
    )
    .action(
      makeAction<
        {
          to?: string;
          firstMatch?: boolean;
          continueOnError?: boolean;
        },
        [string[]]
      >(program, async (deps, g, cmdOpts, messageIds) => {
        const result = await moveMail.run(deps, messageIds, cmdOpts);
        const mode = resolveOutputMode(g);
        if (mode === 'table') {
          emitResult(
            toMoveMailRows(result),
            mode,
            MOVE_MAIL_COLUMNS as unknown as ColumnSpec<unknown>[],
          );
        } else {
          emitResult(result, mode);
        }
        // Partial-failure rule (§10.7 / plan-002 §P5d): any entry in
        // `failed[]` must surface as exit 5 even though `run()` returned
        // normally. The payload is already emitted above.
        if (result.failed.length > 0) {
          process.exitCode = 5;
        }
      }),
    );

  // -------- send-mail --------
  program
    .command('send-mail')
    .description(
      'Send a new email. Default: creates a draft and activates Outlook desktop. ' +
        '--send-now bypasses the draft and sends immediately.',
    )
    .option('--to <recipients...>', 'TO recipients (comma-separated string and/or repeat flag)')
    .option('--cc <recipients...>', 'CC recipients (comma + repeat)')
    .option('--bcc <recipients...>', 'BCC recipients (comma + repeat)')
    .option('--subject <s>', 'Subject line')
    .option('--html <file>', 'HTML body file path')
    .option('--text <file>', 'Plain-text body file path (if no --html)')
    .option(
      '--attach <file>',
      'Attach file (repeatable). Combined size capped at 30 MB.',
      (v: string, acc: string[] = []) => [...acc, v],
      [] as string[],
    )
    .option(
      '--signature <file>',
      'Override signature file path (default: ~/.outlook-cli/signature.html)',
    )
    .option('--no-signature', 'Do not append signature even if signature.html exists')
    .option(
      '--no-cc-self',
      'Suppress automatic CC to authenticated user (CLAUDE.md mandates CC-self by default)',
    )
    .option('--no-save-sent', 'Do not save to SentItems (only meaningful with --send-now)')
    .option('--send-now', 'Send immediately, skip draft + Outlook activation', false)
    .option('--no-open', 'Do not activate Outlook desktop after creating the draft')
    .option('--dry-run', 'Print payload, do not contact M365', false)
    .action(
      makeAction<
        {
          to?: string | string[];
          cc?: string | string[];
          bcc?: string | string[];
          subject?: string;
          html?: string;
          text?: string;
          attach?: string[];
          signature?: string;
          noSignature?: boolean;
          ccSelf?: boolean;
          saveSent?: boolean;
          sendNow?: boolean;
          open?: boolean;
          dryRun?: boolean;
        },
        []
      >(program, async (deps, g, cmdOpts) => {
        const result = await sendMail.run(deps, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- capture-signature --------
  program
    .command('capture-signature')
    .description(
      'Extract email signature from a SentItems message and save to ~/.outlook-cli/signature.html',
    )
    .option('--from-message <id>', 'Source message id (default: latest in SentItems)')
    .option('--out <file>', 'Output path (default: ~/.outlook-cli/signature.html)')
    .action(
      makeAction<{ fromMessage?: string; out?: string }, []>(program, async (deps, g, cmdOpts) => {
        const result = await captureSignature.run(deps, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- reply <id> --------
  program
    .command('reply')
    .argument('<id>', 'Source message id to reply to')
    .description('Reply to a message (draft-first by default; --send-now to dispatch)')
    .option('--html <file>', 'HTML body file (your reply content; auto-quote + signature appended)')
    .option('--text <file>', 'Plain-text body file (escaped + wrapped in <p>)')
    .option(
      '--signature <file>',
      'Override signature file path (default: ~/.outlook-cli/signature.html)',
    )
    .option('--no-signature', 'Suppress signature appending')
    .option(
      '--no-cc-self',
      'Suppress automatic CC to authenticated user (default: ON, per CLAUDE.md compliance)',
    )
    .option('--send-now', 'Send immediately, skip draft + Outlook activation', false)
    .option('--no-open', 'Do not activate Outlook desktop after creating the draft')
    .option('--dry-run', 'Print result without contacting M365', false)
    .action(
      makeAction<
        {
          html?: string;
          text?: string;
          signature?: string;
          noSignature?: boolean;
          ccSelf?: boolean;
          sendNow?: boolean;
          open?: boolean;
          dryRun?: boolean;
        },
        [string]
      >(program, async (deps, g, cmdOpts, id) => {
        const result = await reply.run(deps, 'reply', id, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- reply-all <id> --------
  program
    .command('reply-all')
    .argument('<id>', 'Source message id to reply-all to')
    .description('Reply-all to a message (recipients pre-pop from server)')
    .option('--html <file>', 'HTML body file')
    .option('--text <file>', 'Plain-text body file')
    .option('--signature <file>', 'Override signature file path')
    .option('--no-signature', 'Suppress signature appending')
    .option('--no-cc-self', 'Suppress automatic CC to authenticated user')
    .option('--send-now', 'Send immediately, skip draft + Outlook activation', false)
    .option('--no-open', 'Do not activate Outlook desktop')
    .option('--dry-run', 'Print result without contacting M365', false)
    .action(
      makeAction<
        {
          html?: string;
          text?: string;
          signature?: string;
          noSignature?: boolean;
          ccSelf?: boolean;
          sendNow?: boolean;
          open?: boolean;
          dryRun?: boolean;
        },
        [string]
      >(program, async (deps, g, cmdOpts, id) => {
        const result = await reply.run(deps, 'reply-all', id, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // -------- forward <id> --------
  program
    .command('forward')
    .argument('<id>', 'Source message id to forward')
    .description('Forward a message — auto-quotes original. --to required (forward target).')
    .option('--to <recipients...>', 'Forward target(s) (comma + repeat). REQUIRED.')
    .option('--cc <recipients...>', 'CC on the forward')
    .option('--bcc <recipients...>', 'BCC on the forward')
    .option('--html <file>', 'HTML body file (your forwarding note)')
    .option('--text <file>', 'Plain-text body file')
    .option('--signature <file>', 'Override signature file path')
    .option('--no-signature', 'Suppress signature appending')
    .option('--no-cc-self', 'Suppress automatic CC to authenticated user')
    .option('--send-now', 'Send immediately, skip draft + Outlook activation', false)
    .option('--no-open', 'Do not activate Outlook desktop')
    .option('--dry-run', 'Print result without contacting M365', false)
    .action(
      makeAction<
        {
          to?: string | string[];
          cc?: string | string[];
          bcc?: string | string[];
          html?: string;
          text?: string;
          signature?: string;
          noSignature?: boolean;
          ccSelf?: boolean;
          sendNow?: boolean;
          open?: boolean;
          dryRun?: boolean;
        },
        [string]
      >(program, async (deps, g, cmdOpts, id) => {
        const result = await reply.run(deps, 'forward', id, cmdOpts);
        emitResult(result, resolveOutputMode(g));
      }),
    );

  // Let commander handle its own errors (invalid argv → exit 2).
  program.exitOverride((err) => {
    // commander exits for --help / --version via a CommanderError with a
    // zero-ish `exitCode`. Honour it so the process can terminate cleanly.
    if (err.exitCode === 0) {
      void exitWithDrain(0);
      return;
    }
    throw err;
  });

  try {
    await program.parseAsync(argv);
    const ec = process.exitCode;
    if (typeof ec === 'number') return ec;
    if (typeof ec === 'string') {
      const n = Number.parseInt(ec, 10);
      return Number.isFinite(n) ? n : 0;
    }
    return 0;
  } catch (err) {
    return reportError(err);
  }
}

/**
 * commander option parser for integer-valued flags like `--top`.
 * Passes through invalid input as-is so command-level validators emit
 * command-specific error messages.
 */
function parseIntArg(v: string): number {
  if (!/^-?\d+$/.test(v)) {
    throw new CommanderLikeError(`expected integer value, got ${JSON.stringify(v)}`);
  }
  return Number.parseInt(v, 10);
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/**
 * Exit only after stdout (and stderr) have drained to a piped consumer.
 * `process.exit()` on its own does NOT wait for buffered writes to flush
 * to a downstream pipe — large JSON payloads got truncated at ~64KB on a
 * pipe boundary because the Node runtime tore down before the kernel
 * pipe drained. We yield the event loop until both streams report
 * `writableLength === 0` before calling `process.exit`.
 */
export async function exitWithDrain(code: number): Promise<never> {
  await Promise.all([drainStream(process.stdout), drainStream(process.stderr)]);
  process.exit(code);
}

function drainStream(stream: NodeJS.WriteStream): Promise<void> {
  return new Promise((resolve) => {
    if (stream.writableLength === 0 || stream.writableEnded) {
      resolve();
      return;
    }
    stream.once('drain', () => resolve());
  });
}

// Only invoke main when executed as a script (not when imported by tests).
if (require.main === module) {
  main(process.argv).then(
    (code) => {
      void exitWithDrain(code);
    },
    (err) => {
      const code = reportError(err);
      void exitWithDrain(code);
    },
  );
}
