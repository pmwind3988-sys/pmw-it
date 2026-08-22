import { describe, it, expect } from 'vitest';
import { detectHeader, toGrid } from './detectHeader.js';

describe('detectHeader', () => {
  it('finds a header on the first row', () => {
    const rows = [['Name', 'Amount'], ['a', 1], ['b', 2]];
    expect(detectHeader(rows).headerIndex).toBe(0);
  });

  it('skips a title banner above the header', () => {
    const rows = [
      ['IT Request Report 2024', null],
      [null, null],
      ['Name', 'Amount'],
      ['a', 1],
    ];
    expect(detectHeader(rows).headerIndex).toBe(2);
  });

  it('skips leading blank rows', () => {
    const rows = [[null, null], ['', ''], ['Name', 'Amount'], ['a', 1]];
    expect(detectHeader(rows).headerIndex).toBe(2);
  });

  it('returns -1 when no row looks like a header', () => {
    expect(detectHeader([[1, 2], [3, 4]]).headerIndex).toBe(-1);
  });

  it('returns -1 for an empty sheet', () => {
    expect(detectHeader([]).headerIndex).toBe(-1);
  });
});

describe('toGrid', () => {
  it('splits headers from data rows', () => {
    const rows = [['Name', 'Amount'], ['a', 1]];
    expect(toGrid(rows, 0)).toEqual({ headers: ['Name', 'Amount'], rows: [['a', 1]] });
  });

  it('de-duplicates repeated header names', () => {
    const rows = [['Name', 'Name', 'Name'], ['a', 'b', 'c']];
    expect(toGrid(rows, 0).headers).toEqual(['Name', 'Name (2)', 'Name (3)']);
  });

  it('names blank header cells by position', () => {
    const rows = [['Name', ''], ['a', 'b']];
    expect(toGrid(rows, 0).headers).toEqual(['Name', 'Column 2']);
  });
});
