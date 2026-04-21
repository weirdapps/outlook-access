/**
 * stdout rendering for the Outlook CLI.
 *
 * Two modes:
 *   - 'json'  → `JSON.stringify(data, null, 2)`.
 *   - 'table' → hand-rolled ASCII table (no Unicode box-drawing, no third-party
 *               deps). Columns are separated by two spaces and a dashed
 *               separator line lives between the header and the data rows.
 *
 * Design rationale: the CLI targets plain terminals on macOS, Linux, and
 * Windows (cmd.exe, PowerShell). ASCII-only output avoids the Windows console
 * codepage issues that Unicode box-drawing characters frequently trigger.
 */

export type OutputMode = 'json' | 'table';

export interface ColumnSpec<T> {
  /** Human-readable header for the table column. */
  header: string;
  /** Extract the cell string from a row. */
  extract: (row: T) => string;
  /** Optional cap on rendered cell width; longer cells are ellipsized mid-string. */
  maxWidth?: number;
}

/** Unicode horizontal ellipsis (single char, counts as width 1). */
const ELLIPSIS = '…';

/**
 * Render `data` to a string according to `mode`.
 *
 * In 'table' mode, `columns` is required (throws if omitted). If `data` is not
 * an array it is wrapped as a single-row table.
 */
export function formatOutput<T>(
  data: T | T[],
  mode: OutputMode,
  columns?: ColumnSpec<T>[],
): string {
  if (mode === 'json') {
    return JSON.stringify(data, null, 2);
  }

  if (mode !== 'table') {
    throw new Error(`formatOutput: unsupported mode "${String(mode)}"`);
  }

  if (!columns || columns.length === 0) {
    throw new Error('formatOutput: table mode requires a non-empty columns array');
  }

  const rows: T[] = Array.isArray(data) ? data : [data];

  // Render each cell first — call extract() exactly once per row/column.
  const rendered: string[][] = rows.map((row) =>
    columns.map((col) => {
      const raw = col.extract(row);
      const str = raw === undefined || raw === null ? '' : String(raw);
      return truncate(str, col.maxWidth);
    }),
  );

  // Compute column widths: max(headerWidth, widestCell), capped by maxWidth.
  const widths: number[] = columns.map((col, idx) => {
    let w = col.header.length;
    for (const row of rendered) {
      const cellW = row[idx]!.length;
      if (cellW > w) w = cellW;
    }
    if (typeof col.maxWidth === 'number' && col.maxWidth > 0 && w > col.maxWidth) {
      w = col.maxWidth;
    }
    return w;
  });

  const headerLine = columns.map((col, i) => padRight(col.header, widths[i]!)).join('  ');
  const separatorLine = widths.map((w) => '-'.repeat(w)).join('  ');
  const dataLines = rendered.map((row) =>
    row.map((cell, i) => padRight(cell, widths[i]!)).join('  '),
  );

  return [headerLine, separatorLine, ...dataLines].join('\n');
}

/**
 * Truncate `s` so its visible length is at most `maxWidth`. Ellipsizes in the
 * middle (keeps first half + "…" + last chunk) to preserve useful prefix/suffix
 * context such as filenames like "long-report-...-final.pdf".
 *
 * If `maxWidth` is undefined, non-positive, or larger than the string, the
 * string is returned unchanged.
 */
function truncate(s: string, maxWidth: number | undefined): string {
  if (typeof maxWidth !== 'number' || maxWidth <= 0) return s;
  if (s.length <= maxWidth) return s;
  if (maxWidth === 1) return ELLIPSIS;
  // Split remaining budget (maxWidth - 1 for the ellipsis) between head and tail.
  const budget = maxWidth - 1;
  const head = Math.ceil(budget / 2);
  const tail = budget - head;
  const headPart = s.slice(0, head);
  const tailPart = tail > 0 ? s.slice(s.length - tail) : '';
  return headPart + ELLIPSIS + tailPart;
}

/** Left-align string in a fixed-width column (pad right with spaces). */
function padRight(s: string, width: number): string {
  if (s.length >= width) return s;
  return s + ' '.repeat(width - s.length);
}
