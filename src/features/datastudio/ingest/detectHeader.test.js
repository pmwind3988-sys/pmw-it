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

  // Found by running the real pipeline over a generated workbook. The
  // header was compared against ONE row below it, and that row happened
  // to carry 'pending' in the money column. That single cell erased the
  // strings-over-numbers signal, and the one-cell title banner above won
  // instead -- naming every column after the report title and shifting
  // the whole table up a row.
  it('is not thrown off by a stray text value in the first data row', () => {
    const rows = [
      ['IT Asset Report 2024', null, null],
      ['Asset Tag', 'Department', 'Cost'],
      ['A1000', 'HR', 'pending'],
      ['A1001', 'IT', 1007],
      ['A1002', 'Finance', 1014],
      ['A1003', 'Ops', 1021],
    ];
    expect(detectHeader(rows).headerIndex).toBe(1);
  });

  // The other half of the same bug: a sparse banner must not out-score a
  // full header just because its empty columns differ from a full body.
  it('does not mistake a one-cell title banner for a header', () => {
    const rows = [
      ['Quarterly Summary', null, null, null],
      ['Region', 'Owner', 'Units', 'Value'],
      ['North', 'Ann', 12, 300],
      ['South', 'Bo', 9, 210],
    ];
    expect(detectHeader(rows).headerIndex).toBe(1);
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
