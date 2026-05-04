// src/auth/lock.ts
//
// Unit 3 — Advisory PID lock for the browser-capture flow.
// Design ref: docs/design/project-design.md §2.6
//
// Creates a lock file with O_CREAT|O_EXCL|O_WRONLY ('wx') at mode 0o600.
// The file's content is a single line containing the owner process's PID.
// On EEXIST, we inspect the stored PID: if the process is dead (ESRCH on
// signal 0) we treat the lock as stale, remove it and retry exactly once.

import * as fs from 'node:fs';
import * as nodePath from 'node:path';

/**
 * Acquire an advisory lock at `path`. Returns a release function that is
 * safe to call multiple times (idempotent).
 *
 * @throws Error('another outlook-cli instance holds the lock: ' + path)
 *         when a live process already owns the lock.
 */
export async function acquireLock(path: string): Promise<() => Promise<void>> {
  // Ensure the parent directory exists (mode 0o700 for privacy). First-run
  // case: $HOME/.outlook-cli/ may not exist yet.
  fs.mkdirSync(nodePath.dirname(path), { recursive: true, mode: 0o700 });

  const tryOpen = (): number => fs.openSync(path, 'wx', 0o600);

  let fd: number;
  try {
    fd = tryOpen();
  } catch (err) {
    if (!isEexist(err)) throw err;

    // Lock file exists — inspect its PID.
    const existingPid = readLockPid(path);

    if (existingPid !== null && isProcessAlive(existingPid)) {
      throw new Error('another outlook-cli instance holds the lock: ' + path, {
        cause: err,
      });
    }

    // Stale (or unreadable) lock — remove and retry exactly once.
    try {
      fs.unlinkSync(path);
    } catch (unlinkErr) {
      if (!isEnoent(unlinkErr)) throw unlinkErr;
    }

    try {
      fd = tryOpen();
    } catch (err2) {
      if (isEexist(err2)) {
        throw new Error('another outlook-cli instance holds the lock: ' + path, {
          cause: err2,
        });
      }
      throw err2;
    }
  }

  try {
    fs.writeSync(fd, `${process.pid}\n`);
  } finally {
    try {
      fs.closeSync(fd);
    } catch {
      // Closing a failed write shouldn't mask the primary error path.
    }
  }

  let released = false;
  const release = async (): Promise<void> => {
    if (released) return;
    released = true;
    try {
      fs.unlinkSync(path);
    } catch (err) {
      if (!isEnoent(err)) throw err;
      // Already gone — idempotent success.
    }
  };

  return release;
}

// ── Internals ────────────────────────────────────────────────────────────────

function readLockPid(path: string): number | null {
  let content: string;
  try {
    content = fs.readFileSync(path, 'utf8');
  } catch {
    return null;
  }

  const firstLine = content.split('\n')[0]?.trim() ?? '';
  if (firstLine.length === 0) return null;

  const n = Number.parseInt(firstLine, 10);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n;
}

function isProcessAlive(pid: number): boolean {
  try {
    // Signal 0 — existence / permission probe; never actually delivered.
    process.kill(pid, 0);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException)?.code;
    if (code === 'ESRCH') return false;
    // EPERM means the process exists but belongs to another user — treat as alive.
    if (code === 'EPERM') return true;
    // Anything else: conservatively assume alive to avoid clobbering a real lock.
    return true;
  }
}

function isEexist(err: unknown): boolean {
  return (err as NodeJS.ErrnoException | undefined)?.code === 'EEXIST';
}

function isEnoent(err: unknown): boolean {
  return (err as NodeJS.ErrnoException | undefined)?.code === 'ENOENT';
}
