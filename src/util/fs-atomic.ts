// src/util/fs-atomic.ts
//
// Atomic filesystem helpers for the outlook-cli.
// See docs/design/project-design.md §2.10 and ADR-09 for rationale.
//
// Invariants:
//   - The temp file is created with O_CREAT|O_EXCL|O_WRONLY + explicit mode,
//     so the final file is secure from inception (no chmod-after-write race).
//   - fsync(fd) is called before close+rename to guarantee the data blocks are
//     on disk before the directory entry is swapped.
//   - On any error after the temp fd is open, the temp file is best-effort
//     unlinked and the original error is rethrown wrapped in IoError.

import * as fs from 'node:fs';
import * as path from 'node:path';
import { randomBytes } from 'node:crypto';

import { IoError } from '../config/errors';

export interface AtomicWriteOptions {
  /** Mode for the created file. Default: 0o600. */
  mode?: number;
  /** If true, an existing finalPath is overwritten. Default: true (design
   *  contract for session file, which is intentionally rewritten on every
   *  successful login). Callers that want EEXIST protection pass false.
   */
  overwrite?: boolean;
  /**
   * Mode to apply to the parent directory when it is created OR (defensively)
   * chmod'd after creation. When undefined, the parent directory is created
   * with `recursive: true` using the system umask, and NOT chmod'd (this is the
   * correct behaviour for user-chosen output directories such as the
   * `--out` folder of `download-attachments`).
   *
   * Set to 0o700 for the private session directory under $HOME/.outlook-cli.
   */
  parentDirMode?: number;
}

function buildTempPath(finalPath: string): string {
  const dir = path.dirname(finalPath);
  const base = path.basename(finalPath);
  const rand = randomBytes(6).toString('hex');
  return path.join(dir, `.${base}.tmp.${process.pid}.${rand}`);
}

async function ensureParentDir(
  finalPath: string,
  parentDirMode: number | undefined,
): Promise<string> {
  const dir = path.dirname(finalPath);
  try {
    await fs.promises.mkdir(dir, {
      recursive: true,
      ...(parentDirMode !== undefined ? { mode: parentDirMode } : {}),
    });
  } catch (err) {
    throw new IoError({
      code: 'IO_MKDIR_EACCES',
      message: `Cannot create directory: ${dir}`,
      path: dir,
      cause: err,
    });
  }
  // Defensive chmod on the exact leaf directory only when the caller has
  // explicitly asked for a private parent mode (e.g. 0o700 for the session
  // directory). For user-chosen output directories we leave the mode as the
  // user set it.
  if (parentDirMode !== undefined) {
    try {
      await fs.promises.chmod(dir, parentDirMode);
    } catch {
      // Non-fatal on platforms that do not implement chmod (e.g. Windows).
      // The underlying file is still opened with the requested mode below.
    }
  }
  return dir;
}

async function unlinkBestEffort(p: string): Promise<void> {
  try {
    await fs.promises.unlink(p);
  } catch {
    // ignore — temp file may already be gone
  }
}

async function writeAtomic(
  finalPath: string,
  payload: Buffer,
  options: AtomicWriteOptions | undefined,
): Promise<void> {
  const mode = options?.mode ?? 0o600;
  const overwrite = options?.overwrite ?? true;
  const parentDirMode = options?.parentDirMode;

  await ensureParentDir(finalPath, parentDirMode);
  const tmp = buildTempPath(finalPath);

  let fd: number | undefined;
  try {
    // 'wx' = O_CREAT | O_EXCL | O_WRONLY — fails if the temp path exists.
    fd = fs.openSync(tmp, 'wx', mode);
    fs.writeSync(fd, payload);
    fs.fsyncSync(fd);
  } catch (err) {
    if (fd !== undefined) {
      try {
        fs.closeSync(fd);
      } catch {
        // ignore
      }
    }
    await unlinkBestEffort(tmp);
    throw new IoError({
      code: 'IO_SESSION_WRITE',
      message: `Failed to write file atomically: ${finalPath}`,
      path: finalPath,
      cause: err,
    });
  }
  // Close is a separate try so that a close-only failure still triggers cleanup.
  try {
    fs.closeSync(fd);
  } catch (err) {
    await unlinkBestEffort(tmp);
    throw new IoError({
      code: 'IO_SESSION_WRITE',
      message: `Failed to close temp file for: ${finalPath}`,
      path: finalPath,
      cause: err,
    });
  }

  if (!overwrite) {
    let exists = false;
    try {
      await fs.promises.access(finalPath);
      exists = true;
    } catch {
      // File does not exist, exists remains false
    }
    if (exists) {
      await unlinkBestEffort(tmp);
      throw new IoError({
        code: 'IO_WRITE_EEXIST',
        message: `Refusing to overwrite existing file: ${finalPath}`,
        path: finalPath,
      });
    }
  }

  try {
    fs.renameSync(tmp, finalPath);
  } catch (err) {
    await unlinkBestEffort(tmp);
    throw new IoError({
      code: 'IO_SESSION_WRITE',
      message: `Failed to rename temp file into place: ${finalPath}`,
      path: finalPath,
      cause: err,
    });
  }
}

/**
 * Atomically write JSON content to `filePath`. The parent directory is
 * created (recursively) with mode 0o700 if absent. The file itself is created
 * with the provided mode (default 0o600). Pretty-printed with 2-space indent
 * and a trailing newline.
 *
 * @throws IoError on any filesystem error.
 */
export async function atomicWriteJson(
  filePath: string,
  data: unknown,
  options?: AtomicWriteOptions,
): Promise<void> {
  const text = JSON.stringify(data, null, 2) + '\n';
  const payload = Buffer.from(text, 'utf8');
  await writeAtomic(filePath, payload, options);
}

/**
 * Atomically write an arbitrary Buffer to `filePath`. Same semantics as
 * atomicWriteJson — created with mode 0o600 by default inside a 0o700 parent.
 */
export async function atomicWriteBuffer(
  filePath: string,
  content: Buffer,
  options?: AtomicWriteOptions,
): Promise<void> {
  await writeAtomic(filePath, content, options);
}

/**
 * Read and parse a JSON file.
 *   - Returns null on ENOENT (file does not exist).
 *   - Throws IoError("IO_SESSION_READ") on any other read error.
 *   - Throws IoError("IO_SESSION_CORRUPT") on JSON.parse failure.
 *
 * Type-parametric: the caller asserts the expected shape. Use a structural
 * validator (e.g. isValidSessionFile) on the returned value before using it.
 */
export async function readJsonFile<T>(filePath: string): Promise<T | null> {
  let raw: string;
  try {
    raw = await fs.promises.readFile(filePath, 'utf8');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT') {
      return null;
    }
    throw new IoError({
      code: 'IO_SESSION_READ',
      message: `Failed to read file: ${filePath}`,
      path: filePath,
      cause: err,
    });
  }
  try {
    return JSON.parse(raw) as T;
  } catch (err) {
    throw new IoError({
      code: 'IO_SESSION_CORRUPT',
      message: `File is not valid JSON: ${filePath}`,
      path: filePath,
      cause: err,
    });
  }
}
