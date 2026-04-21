// src/session/store.ts
//
// Atomic persistence for the outlook-cli session file.
// See docs/design/project-design.md §2.4 and refined-request-outlook-cli.md §7.

import * as fs from 'node:fs';
import * as path from 'node:path';

import { IoError } from '../config/errors';
import { atomicWriteJson, readJsonFile } from '../util/fs-atomic';
import { SessionFile, isValidSessionFile } from './schema';

/** 60-second grace window applied when judging token expiry. */
export const EXPIRY_SKEW_MS = 60_000;

/**
 * Read a session file if present. Returns null when:
 *   - the file does not exist (ENOENT), or
 *   - the file contents fail the structural schema check (a warning is
 *     written to stderr so the caller can tell something was found but
 *     discarded; the secret material is never logged).
 *
 * Throws IoError on filesystem errors other than ENOENT.
 */
export async function loadSession(
  filePath: string,
): Promise<SessionFile | null> {
  const parsed = await readJsonFile<unknown>(filePath);
  if (parsed === null) {
    return null;
  }
  if (!isValidSessionFile(parsed)) {
    // Redaction: we never dump the parsed payload — it may include a bearer
    // token or cookie values. Only the path and a generic message are logged.
    process.stderr.write(
      `warning: session file at ${filePath} failed schema validation; ignoring.\n`,
    );
    return null;
  }
  return parsed;
}

/**
 * Persist the session file atomically with mode 0600 inside a 0700 parent dir.
 * Delegates to atomicWriteJson — see src/util/fs-atomic.ts for the exact
 * create-temp / fsync / rename algorithm.
 *
 * @throws IoError("IO_SESSION_WRITE") on any filesystem error.
 */
export async function saveSession(
  filePath: string,
  s: SessionFile,
): Promise<void> {
  // atomicWriteJson (with parentDirMode: 0o700) creates the parent directory
  // with mode 0o700 and defensively chmods it. We additionally pre-create the
  // directory here so that a later chmod-only failure (e.g. on Windows) still
  // leaves a usable session dir.
  const dir = path.dirname(filePath);
  try {
    await fs.promises.mkdir(dir, { recursive: true, mode: 0o700 });
    try {
      await fs.promises.chmod(dir, 0o700);
    } catch {
      // Non-fatal (e.g. Windows).
    }
  } catch (err) {
    throw new IoError({
      code: 'IO_MKDIR_EACCES',
      message: `Cannot create session directory: ${dir}`,
      path: dir,
      cause: err,
    });
  }
  await atomicWriteJson(filePath, s, {
    mode: 0o600,
    overwrite: true,
    parentDirMode: 0o700,
  });
}

/**
 * Return true when the bearer token is considered expired.
 *
 * A 60-second safety skew is applied so we refresh proactively rather than
 * racing the server: if `nowMs + 60_000 >= bearer.expiresAt`, we say expired.
 *
 * If `expiresAt` is unparsable, the session is treated as expired (safe default).
 */
export function isExpired(s: SessionFile, nowMs?: number): boolean {
  const now = typeof nowMs === 'number' ? nowMs : Date.now();
  const expiresMs = Date.parse(s.bearer.expiresAt);
  if (!Number.isFinite(expiresMs)) {
    return true;
  }
  return now + EXPIRY_SKEW_MS >= expiresMs;
}

/**
 * Delete the session file if it exists. No-op if it does not.
 * Used by tests and by `login --force`.
 *
 * @throws IoError on filesystem errors other than ENOENT.
 */
export async function deleteSession(filePath: string): Promise<void> {
  try {
    await fs.promises.unlink(filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT') {
      return;
    }
    throw new IoError({
      code: 'IO_SESSION_WRITE',
      message: `Failed to delete session file: ${filePath}`,
      path: filePath,
      cause: err,
    });
  }
}
