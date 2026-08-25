import { describe, it, expect } from 'vitest';
import { buildDataset } from './dataset.js';
import { buildMask, maskFor, createMaskCache } from './filterMask.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';

const grid = {
  headers: ['Dept', 'Amount'],
  rows: [['HR', '10'], ['IT', '20'], ['HR', '30'], ['Finance', '40']],
};
const ds = applyCleanPlan(grid, [], profileDataset(grid));
const count = (m) => m.reduce((a, b) => a + b, 0);

describe('buildMask', () => {
  it('keeps every row when there are no filters', () => {
    expect(count(buildMask(ds, []))).toBe(4);
  });

  it('filters a categorical column by membership', () => {
    const m = buildMask(ds, [{ column: 'Dept', kind: 'in', values: ['HR'] }]);
    expect(count(m)).toBe(2);
    expect(Array.from(m)).toEqual([1, 0, 1, 0]);
  });

  it('accepts multiple values on one filter', () => {
    const m = buildMask(ds, [{ column: 'Dept', kind: 'in', values: ['HR', 'IT'] }]);
    expect(count(m)).toBe(3);
  });

  it('ANDs separate filters together', () => {
    const m = buildMask(ds, [
      { column: 'Dept', kind: 'in', values: ['HR'] },
      { column: 'Amount', kind: 'range', min: 20, max: 100 },
    ]);
    expect(count(m)).toBe(1);
  });

  it('excludes NaN from range filters', () => {
    // RULING (amends the plan's fixture, not its assertion): the plan
    // gave this column ['1', 'nope', '3'], which is 2/3 numeric and so
    // types CATEGORICAL under spec §7.2's 95% rule -- a range filter on
    // it could never have kept 2 rows. Widened to twenty values (19
    // numeric is exactly 95%), of which exactly two fall in 0..10 and
    // one is the same unparseable 'nope'. The assertion is the plan's.
    const withNull = {
      headers: ['n'],
      rows: [['1'], ['nope'], ['3'], ...Array.from({ length: 17 }, () => ['100'])],
    };
    const d2 = applyCleanPlan(withNull, [], profileDataset(withNull));
    expect(count(buildMask(d2, [{ column: 'n', kind: 'range', min: 0, max: 10 }])))
      .toBe(2);
  });

  it('treats range bounds as inclusive at both ends', () => {
    const m = buildMask(ds, [{ column: 'Amount', kind: 'range', min: 20, max: 30 }]);
    expect(count(m)).toBe(2);
  });

  it('drops rows whose category is missing when filtering by membership', () => {
    const sparse = {
      headers: ['Dept', 'Site'],
      rows: [['HR', 'a'], [null, 'b'], ['IT', 'c']],
    };
    const d2 = applyCleanPlan(sparse, [], profileDataset(sparse));
    const m = buildMask(d2, [{ column: 'Dept', kind: 'in', values: ['HR', 'IT'] }]);
    expect(Array.from(m)).toEqual([1, 0, 1]);
  });

  // A filter can name a column this dataset does not have.
  // Ignoring the filter shows more than asked; treating it as matching
  // nothing shows an empty dashboard with no explanation. Ignoring is
  // the lesser harm, and the filter bar still lists what is in force.
  it('ignores a filter naming a column that is not in the dataset', () => {
    expect(count(buildMask(ds, [{ column: 'Gone', kind: 'in', values: ['x'] }]))).toBe(4);
  });

  it('ignores a membership filter with no values, rather than hiding everything', () => {
    expect(count(buildMask(ds, [{ column: 'Dept', kind: 'in', values: [] }]))).toBe(4);
  });
});

describe('maskFor -- the self-exclusion rule (spec §10.3)', () => {
  const selection = { sourceTileId: 'tile_1', column: 'Dept', values: ['HR'] };

  it('does not filter the tile that originated the selection', () => {
    expect(count(maskFor(ds, [], selection, 'tile_1'))).toBe(4);
  });

  it('filters every other tile', () => {
    expect(count(maskFor(ds, [], selection, 'tile_2'))).toBe(2);
  });

  it('still applies global filters to the source tile', () => {
    const globals = [{ column: 'Amount', kind: 'range', min: 0, max: 25 }];
    expect(count(maskFor(ds, globals, selection, 'tile_1'))).toBe(2);
  });

  it('behaves normally when there is no selection', () => {
    expect(count(maskFor(ds, [], null, 'tile_1'))).toBe(4);
  });
});

describe('createMaskCache', () => {
  it('returns the identical array for identical inputs', () => {
    const cache = createMaskCache();
    const a = cache.get(ds, [], null, 'tile_1');
    const b = cache.get(ds, [], null, 'tile_1');
    expect(a).toBe(b);
  });

  it('shares one array between tiles that are not the selection source', () => {
    const cache = createMaskCache();
    const sel = { sourceTileId: 'tile_1', column: 'Dept', values: ['HR'] };
    expect(cache.get(ds, [], sel, 'tile_2')).toBe(cache.get(ds, [], sel, 'tile_3'));
  });

  it('returns a different array once the filters change', () => {
    const cache = createMaskCache();
    const a = cache.get(ds, [], null, 'tile_1');
    const b = cache.get(ds, [{ column: 'Dept', kind: 'in', values: ['HR'] }], null, 'tile_1');
    expect(a).not.toBe(b);
  });

  // The source tile and the rest ask different questions, so they must
  // not collide in the cache -- that would be the self-exclusion rule
  // silently undone by a memo.
  it('gives the source tile a different array from the others', () => {
    const cache = createMaskCache();
    const sel = { sourceTileId: 'tile_1', column: 'Dept', values: ['HR'] };
    expect(cache.get(ds, [], sel, 'tile_1')).not.toBe(cache.get(ds, [], sel, 'tile_2'));
  });

  it('does not serve one dataset a mask built for another', () => {
    const cache = createMaskCache();
    const other = applyCleanPlan(
      { headers: ['Dept', 'Amount'], rows: [['HR', '10']] },
      [],
      profileDataset({ headers: ['Dept', 'Amount'], rows: [['HR', '10']] }),
    );
    const a = cache.get(ds, [], null, 'tile_1');
    const b = cache.get(other, [], null, 'tile_1');
    expect(a).not.toBe(b);
    expect(b.length).toBe(1);
  });
});

describe('filtering a multi column', () => {
  const multiDataset = buildDataset({
    headers: ['Challenges'],
    columns: [['A;B;', 'B;', 'C;', '']],
    profile: { columns: [{ name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' }] },
  });

  it('keeps a row when any of its options match', () => {
    const mask = buildMask(multiDataset, [{ column: 'Challenges', kind: 'in', values: ['A'] }]);
    expect(Array.from(mask)).toEqual([1, 0, 0, 0]);
  });

  it('drops a row with no options at all', () => {
    const mask = buildMask(multiDataset, [{ column: 'Challenges', kind: 'in', values: ['A', 'B', 'C'] }]);
    expect(Array.from(mask)).toEqual([1, 1, 1, 0]);
  });
});
