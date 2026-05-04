import path from 'node:path';

/**
 * Windows reserved device names (case-insensitive, any extension).
 * Matches: CON, PRN, AUX, NUL, COM1..COM9, LPT1..LPT9 — optionally followed
 * by an extension. A reserved name matches whether or not it has an extension.
 */
export const WINDOWS_RESERVED: RegExp = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)/i;

/**
 * Characters illegal in filenames on Windows or POSIX:
 *   path separators  : / \
 *   Windows-forbidden: < > : " | ? *
 *   ASCII control    : \x00-\x1F  (also strip DEL \x7F below)
 */
// eslint-disable-next-line no-control-regex -- intentionally matches control chars for filename sanitization
export const ILLEGAL_CHARS: RegExp = /[/\\:*?"<>|\x00-\x1F\x7F]/g;

/**
 * Maximum byte-length of a filename component (reserve 12 bytes for dedup suffix
 * such as " (999).ext"). 255-byte limit is common across ext4/APFS/NTFS.
 */
export const MAX_FILENAME_BYTES: number = 243;

/** Fallback name used when sanitization produces an empty/invalid string. */
const DEFAULT_FALLBACK = 'attachment';

/**
 * Sanitize an arbitrary attachment Name field into a safe filesystem filename.
 *
 * Algorithm (ported from docs/research/outlook-v2-attachments.md §5.1 with
 * minor tweaks — see notes below):
 *
 *   1. NFC Unicode normalization (prevent lookalike-codepoint evasion).
 *   2. Strip path separators, the literal token ".." and ASCII control chars.
 *   3. Replace Windows-illegal chars `<>:"|?*` with `_`.
 *   4. Strip leading dots; trim trailing dots and spaces.
 *   5. Reject Windows reserved device names (CON/PRN/AUX/NUL/COM1..9/LPT1..9)
 *      by prefixing the name with `_`.
 *   6. Truncate to MAX_FILENAME_BYTES (UTF-8), preserving the extension when
 *      possible.
 *   7. If the result is empty or just ".", return "attachment".
 */
export function sanitizeAttachmentName(raw: string): string {
  // Defensive coercion — spec says string but callers may pass null/undefined.
  let name = (raw ?? '').toString();

  // 1. NFC normalization.
  name = name.normalize('NFC');

  // 2. Remove occurrences of ".." (path traversal) before other replacements,
  //    then strip path separators and control chars explicitly. ILLEGAL_CHARS
  //    already covers `/`, `\`, and control chars, but we also want to remove
  //    the `..` sequence as a token (not just replace the dots).
  name = name.split('..').join('');

  // 3. Replace Windows-illegal chars (<>:"|?*) as well as path separators and
  //    control chars with '_'. We use a single pass via ILLEGAL_CHARS so every
  //    disallowed byte becomes an underscore.
  name = name.replace(ILLEGAL_CHARS, '_');

  // 4. Strip leading dots (would create hidden files on POSIX); strip trailing
  //    dots and spaces (illegal on Windows). Use lazy quantifiers to prevent ReDoS.
  name = name.replace(/^\.+?/, '').replace(/[\s.]+?$/, '');

  // 5. Reject Windows reserved device names by prefixing with '_'.
  if (WINDOWS_RESERVED.test(name)) {
    name = `_${name}`;
  }

  // 6. Enforce UTF-8 byte limit, preserving the extension when possible.
  const encoder = new TextEncoder();
  if (encoder.encode(name).byteLength > MAX_FILENAME_BYTES) {
    const ext = path.extname(name);
    const extBytes = encoder.encode(ext).byteLength;
    // If the extension alone already exceeds the limit, drop it entirely and
    // truncate the full string byte-wise.
    if (extBytes >= MAX_FILENAME_BYTES) {
      name = truncateToBytes(name, MAX_FILENAME_BYTES);
    } else {
      const baseName = name.slice(0, name.length - ext.length);
      const base = truncateToBytes(baseName, MAX_FILENAME_BYTES - extBytes);
      name = base + ext;
    }
  }

  // 7. Final emptiness guard.
  if (name.length === 0 || name === '.') {
    return DEFAULT_FALLBACK;
  }

  return name;
}

/**
 * Truncate a string so its UTF-8 encoding fits within `maxBytes`. Preserves
 * whole UTF-16 code units (and therefore whole code points for BMP chars);
 * for surrogate pairs we avoid splitting a pair in half.
 */
function truncateToBytes(s: string, maxBytes: number): string {
  const encoder = new TextEncoder();
  if (encoder.encode(s).byteLength <= maxBytes) return s;
  let lo = 0;
  let hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >>> 1;
    if (encoder.encode(s.slice(0, mid)).byteLength <= maxBytes) {
      lo = mid;
    } else {
      hi = mid - 1;
    }
  }
  // Avoid slicing in the middle of a surrogate pair.
  let cut = lo;
  if (cut > 0 && cut < s.length) {
    const prev = s.charCodeAt(cut - 1);
    if (prev >= 0xd800 && prev <= 0xdbff) {
      cut -= 1;
    }
  }
  return s.slice(0, cut);
}

/**
 * If `desiredName` is not already in `existingNames`, return it unchanged.
 * Otherwise, append " (1)", " (2)", ... before the extension until unique.
 * The caller is responsible for tracking the set — this function is pure.
 */
export function deduplicateFilename(desiredName: string, existingNames: Set<string>): string {
  if (!existingNames.has(desiredName)) {
    return desiredName;
  }

  const ext = path.extname(desiredName);
  const base = desiredName.slice(0, desiredName.length - ext.length);

  let counter = 1;
  // Practical upper bound — avoid unbounded loops if the caller misuses the set.
  const MAX_ATTEMPTS = 10_000;
  while (counter <= MAX_ATTEMPTS) {
    const candidate = `${base} (${counter})${ext}`;
    if (!existingNames.has(candidate)) {
      return candidate;
    }
    counter += 1;
  }
  throw new Error(`deduplicateFilename: exhausted attempts for "${desiredName}"`);
}

/**
 * Defense-in-depth path-traversal guard. Resolves `path.join(baseDir, filename)`
 * and verifies the result is strictly inside the resolved `baseDir`. Throws
 * `Error('path traversal attempt')` if the resolved path escapes the directory.
 *
 * Returns the absolute resolved path on success.
 */
export function assertWithinDir(baseDir: string, filename: string): string {
  const resolvedBase = path.resolve(baseDir);
  const resolvedCandidate = path.resolve(path.join(baseDir, filename));
  const prefix = resolvedBase.endsWith(path.sep) ? resolvedBase : resolvedBase + path.sep;
  if (resolvedCandidate !== resolvedBase && !resolvedCandidate.startsWith(prefix)) {
    throw new Error('path traversal attempt');
  }
  return resolvedCandidate;
}
