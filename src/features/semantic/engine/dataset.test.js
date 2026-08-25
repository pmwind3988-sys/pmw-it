import { describe, it, expect } from 'vitest';
import { buildDataset } from './dataset.js';

const profileOf = (columns) => ({ columns });

describe('buildDataset with a multi column', () => {
  const grid = {
    headers: ['Dept', 'Challenges'],
    columns: [
      ['IT', 'Finance', 'Logistics'],
      ['A;B;', 'B;', ''],
    ],
    profile: profileOf([
      { name: 'Dept', type: 'categorical', role: 'dimension' },
      { name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' },
    ]),
  };

  it('stores option codes with row offsets', () => {
    const dataset = buildDataset(grid);
    const column = dataset.columns[1];

    expect(column.type).toBe('multi');
    expect(column.dictionary).toEqual(['A', 'B']);
    expect(Array.from(column.values)).toEqual([0, 1, 1]);
    expect(Array.from(column.offsets)).toEqual([0, 2, 3, 3]);
  });

  it('keeps rowCount the number of rows, not the number of options', () => {
    // The flat option array is longer than the grid. Deriving rowCount
    // from the first column's values would report 3 here by luck and 5
    // if the multi column happened to come first.
    const dataset = buildDataset({
      headers: ['Challenges', 'Dept'],
      columns: [grid.columns[1], grid.columns[0]],
      profile: profileOf([grid.profile.columns[1], grid.profile.columns[0]]),
    });
    expect(dataset.rowCount).toBe(3);
  });
});
