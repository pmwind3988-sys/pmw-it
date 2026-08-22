import { describe, it, expect } from 'vitest';
import { buildDataset } from './dataset.js';
import { formatCell } from './formatCell.js';
import {
  countRows, pageRowIndexes, stepRow, tableColumns, readRecord, readRows,
} from './rows.js';

const profile = {
  columns: [
    { name: 'Dept', type: 'categorical', role: 'dimension' },
    { name: 'Amount', type: 'numeric', role: 'measure' },
    { name: 'Tools', type: 'multi', role: 'dimension', separator: ';' },
    { name: 'Active', type: 'boolean', role: 'dimension' },
    { name: 'Note', type: 'text', role: 'ignored' },
  ],
};

const ds = buildDataset({
  headers: ['Dept', 'Amount', 'Tools', 'Active', 'Note'],
  columns: [
    ['HR', 'IT', 'HR', null],
    [10, 20, null, 40],
    ['Teams;Excel', 'Excel', '', 'Teams'],
    ['yes', 'no', null, 'yes'],
    ['first', 'second', 'third', 'fourth'],
  ],
  profile,
});

const col = (name) => ds.columns[ds.byName.get(name)];
const mask = Uint8Array.from([1, 0, 1, 1]);

describe('formatCell', () => {
  it('joins a multi column\'s options instead of reading its flat array', () => {
    // The trap: a multi column carries a dictionary, so a decoder that
    // checks for one first reads offsets as codes and prints a real
    // label from the wrong row.
    expect(formatCell(col('Tools'), 0)).toBe('Teams, Excel');
    expect(formatCell(col('Tools'), 1)).toBe('Excel');
  });

  it('leaves a row with no options blank rather than picking one', () => {
    expect(formatCell(col('Tools'), 2)).toBe('');
  });

  it('never invents a value for a missing cell', () => {
    expect(formatCell(col('Dept'), 3)).toBe('');
    expect(formatCell(col('Amount'), 2)).toBe('');
    expect(formatCell(col('Active'), 2)).toBe('');
  });

  it('reads booleans as words', () => {
    expect(formatCell(col('Active'), 0)).toBe('Yes');
    expect(formatCell(col('Active'), 1)).toBe('No');
  });

  it('shows a ratio as a percentage on screen but as a number in a CSV', () => {
    const ratio = { type: 'numeric', values: [0.125], isPercent: true, dictionary: null };
    expect(formatCell(ratio, 0)).toBe('12.5%');
    expect(formatCell(ratio, 0, { percentAsText: false })).toBe('0.125');
  });

  it('renders a date-only column without a time of day', () => {
    const day = {
      type: 'date',
      values: Float64Array.from([Date.UTC(2026, 0, 2)]),
      dictionary: null,
      dateOnly: true,
    };
    expect(formatCell(day, 0)).not.toMatch(/:/);
  });
});

describe('countRows', () => {
  it('counts what the mask keeps', () => {
    expect(countRows(ds, mask)).toBe(3);
  });

  it('counts everything when there is no mask', () => {
    expect(countRows(ds, null)).toBe(4);
  });
});

describe('pageRowIndexes', () => {
  it('returns dataset indexes, not page positions', () => {
    expect(pageRowIndexes(ds, mask, 0, 10)).toEqual([0, 2, 3]);
  });

  it('pages within the filtered set', () => {
    expect(pageRowIndexes(ds, mask, 1, 1)).toEqual([2]);
    expect(pageRowIndexes(ds, mask, 3, 10)).toEqual([]);
  });

  it('stops as soon as the page is full', () => {
    expect(pageRowIndexes(ds, null, 0, 2)).toEqual([0, 1]);
  });
});

describe('stepRow', () => {
  it('skips rows the filters exclude', () => {
    expect(stepRow(ds, mask, 0, 1)).toBe(2);
    expect(stepRow(ds, mask, 2, -1)).toBe(0);
  });

  it('returns null at either end', () => {
    expect(stepRow(ds, mask, 3, 1)).toBe(null);
    expect(stepRow(ds, mask, 0, -1)).toBe(null);
  });
});

describe('tableColumns', () => {
  it('sorts the parked bookkeeping columns last without dropping them', () => {
    const names = tableColumns(ds, null).map((c) => c.name);
    expect(names).toEqual(['Dept', 'Amount', 'Tools', 'Active', 'Note']);
    expect(tableColumns(ds, 2).map((c) => c.name)).toEqual(['Dept', 'Amount']);
  });
});

describe('readRecord', () => {
  it('returns every column, parked ones included and flagged', () => {
    const record = readRecord(ds, 0);
    expect(record.map((f) => f.name)).toHaveLength(5);
    expect(record.find((f) => f.name === 'Note')).toMatchObject({
      parked: true, text: 'first',
    });
  });

  it('marks a blank field rather than hiding it', () => {
    expect(readRecord(ds, 3).find((f) => f.name === 'Dept')).toMatchObject({
      empty: true, text: '',
    });
  });

  it('answers with nothing for a row that is not there', () => {
    expect(readRecord(ds, 99)).toEqual([]);
    expect(readRecord(null, 0)).toEqual([]);
  });
});

describe('readRows', () => {
  it('decodes one page against the chosen columns', () => {
    const rows = readRows(ds, mask, tableColumns(ds, 3), 0, 2);
    expect(rows).toEqual([
      { index: 0, cells: ['HR', '10', 'Teams, Excel'] },
      { index: 2, cells: ['HR', '', ''] },
    ]);
  });
});
