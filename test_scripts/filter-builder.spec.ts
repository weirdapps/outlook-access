// test_scripts/filter-builder.test.ts
import { describe, it, expect } from 'vitest';
import { buildReceivedDateFilter, FilterError } from '../src/http/filter-builder';

describe('buildReceivedDateFilter', () => {
  it('returns empty string when neither bound is set', () => {
    expect(buildReceivedDateFilter(undefined, undefined)).toBe('');
  });

  it('builds a >= clause for since only', () => {
    expect(buildReceivedDateFilter('2026-04-22T07:00:00Z', undefined)).toBe(
      'ReceivedDateTime ge 2026-04-22T07:00:00Z',
    );
  });

  it('builds a < clause for until only', () => {
    expect(buildReceivedDateFilter(undefined, '2026-04-23T00:00:00Z')).toBe(
      'ReceivedDateTime lt 2026-04-23T00:00:00Z',
    );
  });

  it('combines both bounds with and', () => {
    expect(buildReceivedDateFilter('2026-04-22T07:00:00Z', '2026-04-23T00:00:00Z')).toBe(
      'ReceivedDateTime ge 2026-04-22T07:00:00Z and ReceivedDateTime lt 2026-04-23T00:00:00Z',
    );
  });

  it('accepts fractional seconds', () => {
    expect(buildReceivedDateFilter('2026-04-22T07:00:00.123Z', undefined)).toBe(
      'ReceivedDateTime ge 2026-04-22T07:00:00.123Z',
    );
  });

  it('throws FilterError on malformed since', () => {
    expect(() => buildReceivedDateFilter('not-a-date', undefined)).toThrow(FilterError);
  });

  it('throws FilterError on missing Z suffix', () => {
    expect(() => buildReceivedDateFilter('2026-04-22T07:00:00', undefined)).toThrow(FilterError);
  });

  it('throws FilterError when since >= until', () => {
    expect(() => buildReceivedDateFilter('2026-04-23T00:00:00Z', '2026-04-22T00:00:00Z')).toThrow(
      FilterError,
    );
  });

  it('throws FilterError when since equals until', () => {
    expect(() => buildReceivedDateFilter('2026-04-22T00:00:00Z', '2026-04-22T00:00:00Z')).toThrow(
      FilterError,
    );
  });
});
