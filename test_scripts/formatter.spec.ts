// test_scripts/formatter.spec.ts
//
// Unit tests for src/output/formatter.ts — formatOutput.

import { describe, expect, it } from 'vitest';

import { formatOutput, ColumnSpec } from '../src/output/formatter';

describe('formatOutput json mode', () => {
  it('returns JSON.stringify with 2-space indent for an array', () => {
    const data = [{ a: 1 }, { a: 2 }];
    expect(formatOutput(data, 'json')).toBe(JSON.stringify(data, null, 2));
  });

  it('returns JSON.stringify with 2-space indent for a single object', () => {
    const data = { greeting: 'hi', count: 3 };
    expect(formatOutput(data, 'json')).toBe(JSON.stringify(data, null, 2));
  });
});

describe('formatOutput table mode', () => {
  it('throws when columns is omitted in table mode', () => {
    expect(() => formatOutput([{ a: 1 }], 'table')).toThrowError(/table mode requires/);
  });

  it('throws when columns is an empty array', () => {
    expect(() => formatOutput([{ a: 1 }], 'table', [])).toThrowError(/table mode requires/);
  });

  it('renders a header line, separator, and data rows', () => {
    type Row = { id: string; name: string };
    const columns: ColumnSpec<Row>[] = [
      { header: 'ID', extract: (r) => r.id },
      { header: 'Name', extract: (r) => r.name },
    ];
    const data: Row[] = [
      { id: '1', name: 'Alice' },
      { id: '22', name: 'Bob' },
    ];
    const out = formatOutput(data, 'table', columns);
    const lines = out.split('\n');
    expect(lines).toHaveLength(4); // header + separator + 2 rows
    expect(lines[0]).toContain('ID');
    expect(lines[0]).toContain('Name');
    expect(lines[1]).toMatch(/^-+\s+-+$/);
    expect(lines[2]).toContain('Alice');
    expect(lines[3]).toContain('Bob');
  });

  it('ellipsizes values exceeding maxWidth', () => {
    type Row = { long: string };
    const columns: ColumnSpec<Row>[] = [{ header: 'long', extract: (r) => r.long, maxWidth: 10 }];
    const data: Row[] = [{ long: 'abcdefghijklmnopqrstuvwxyz' }];
    const out = formatOutput(data, 'table', columns);
    // One data line; ensure it contains the ellipsis char and no longer cell.
    const dataLine = out.split('\n')[2];
    expect(dataLine).toContain('…');
    // Trim right-padding and confirm width <= maxWidth.
    const trimmed = dataLine.replace(/\s+$/, '');
    expect([...trimmed].length).toBeLessThanOrEqual(10);
  });

  it('wraps a single object as a single-row table', () => {
    type Row = { status: string };
    const columns: ColumnSpec<Row>[] = [{ header: 'status', extract: (r) => r.status }];
    const out = formatOutput({ status: 'ok' }, 'table', columns);
    const lines = out.split('\n');
    expect(lines).toHaveLength(3); // header + separator + 1 row
    expect(lines[2].trim()).toBe('ok');
  });

  it('renders null/undefined cell values as empty string', () => {
    type Row = { v: string | null | undefined };
    const columns: ColumnSpec<Row>[] = [{ header: 'v', extract: (r) => r.v as string }];
    const out = formatOutput([{ v: null }, { v: undefined }], 'table', columns);
    const lines = out.split('\n');
    expect(lines[2].trim()).toBe('');
    expect(lines[3].trim()).toBe('');
  });
});

describe('formatOutput unsupported mode', () => {
  it('throws on an unknown mode', () => {
    expect(() => formatOutput([{ a: 1 }], 'xml' as unknown as 'json')).toThrowError(
      /unsupported mode/,
    );
  });
});
