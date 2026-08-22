import { describe, it, expect } from 'vitest';
import { buildDataset } from './dataset.js';
import {
  aggregate, truncateDate, chooseTruncation, binNumeric,
} from './aggregate.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';
import { profileColumn } from '../profile/profileColumn.js';
import { proposeCleanPlan } from '../clean/proposeCleanPlan.js';

const grid = {
  headers: ['Dept', 'Entity', 'Amount'],
  rows: [
    ['HR', 'pmw', '10'], ['IT', 'pmw', '20'], ['HR', 'pmw-ss', '30'],
    ['IT', 'pmw', '40'], ['HR', 'pmw', 'nope'],
  ],
};

// RULING: Amount is 4 numbers out of 5, which spec §7.2's 95% rule types
// CATEGORICAL -- so every assertion below about summing it would fail
// against the plan's own rules. Rather than widen the fixture (which
// would break the plan's exact expected values of 60/40, HR:3, avg 20),
// the column is forced to numeric through the SAME override the profile
// panel gives the user. Charting a column the inference declined is a
// real, supported path, so the fixture stays exactly as the plan wrote
// it and every assertion keeps its meaning.
const profile = profileDataset(grid);
profile.columns[2] = profileColumn(
  grid.rows.map((row) => row[2]), 'Amount', 2, { type: 'numeric' },
);

const ds = applyCleanPlan(grid, [], profile);
const ALL = new Uint8Array(5).fill(1);

const spec = (over = {}) => ({
  chart: 'bar',
  encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'sum' }], series: null },
  sort: { by: 'y', dir: 'desc' },
  limit: 20,
  ...over,
});

describe('aggregate', () => {
  it('groups by a dimension and sums a measure', () => {
    const r = aggregate(ds, ALL, spec());
    expect(r.categories).toEqual(['IT', 'HR']);
    expect(r.series[0].data).toEqual([60, 40]);
  });

  it('skips NaN when summing rather than producing NaN', () => {
    const r = aggregate(ds, ALL, spec());
    expect(r.series[0].data.every(Number.isFinite)).toBe(true);
  });

  it('counts rows including those with a null measure', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'count' }], series: null },
    }));
    expect(Object.fromEntries(r.categories.map((c, i) => [c, r.series[0].data[i]])))
      .toMatchObject({ HR: 3, IT: 2 });
  });

  it('averages over non-null values only', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'avg' }], series: null },
    }));
    const byCat = Object.fromEntries(r.categories.map((c, i) => [c, r.series[0].data[i]]));
    expect(byCat.HR).toBe(20); // (10 + 30) / 2, not / 3
  });

  it('honours the row mask', () => {
    const mask = new Uint8Array([1, 0, 1, 0, 0]);
    const r = aggregate(ds, mask, spec());
    expect(r.categories).toEqual(['HR']);
    expect(r.series[0].data).toEqual([40]);
  });

  it('splits into one series per series-column value', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: {
        x: { column: 'Dept' },
        y: [{ column: 'Amount', agg: 'sum' }],
        series: { column: 'Entity' },
      },
    }));
    expect(r.series.map((s) => s.name).sort()).toEqual(['pmw', 'pmw-ss']);
    expect(r.series[0].data).toHaveLength(r.categories.length);
  });

  it('pads absent series/category combinations with 0', () => {
    const r = aggregate(ds, ALL, spec({
      encoding: {
        x: { column: 'Dept' },
        y: [{ column: 'Amount', agg: 'sum' }],
        series: { column: 'Entity' },
      },
    }));
    const ss = r.series.find((s) => s.name === 'pmw-ss');
    expect(ss.data).toContain(0);
  });

  it('applies top-N and folds the remainder into Other', () => {
    const r = aggregate(ds, ALL, spec({ limit: 1 }));
    expect(r.categories).toEqual(['IT', 'Other']);
    expect(r.series[0].data).toEqual([60, 40]);
  });

  it('sorts ascending when told to', () => {
    const r = aggregate(ds, ALL, spec({ sort: { by: 'y', dir: 'asc' } }));
    expect(r.categories).toEqual(['HR', 'IT']);
  });

  it('returns empty results for an all-zero mask rather than throwing', () => {
    const r = aggregate(ds, new Uint8Array(5), spec());
    expect(r.categories).toEqual([]);
    expect(r.series[0].data).toEqual([]);
  });

  // 'Other' is a summary of everything that did not make the cut, so it
  // belongs at the end whatever the sort says. Sorted into the middle by
  // its own value it reads as just another category.
  it('keeps Other last even when sorting ascending', () => {
    const r = aggregate(ds, ALL, spec({ limit: 1, sort: { by: 'y', dir: 'asc' } }));
    expect(r.categories[r.categories.length - 1]).toBe('Other');
  });

  it('reports the other aggregations over non-null values', () => {
    const agg = (name) => aggregate(ds, ALL, spec({
      encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: name }], series: null },
      sort: { by: 'x', dir: 'asc' },
    }));
    const byCat = (r) => Object.fromEntries(r.categories.map((c, i) => [c, r.series[0].data[i]]));
    expect(byCat(agg('min')).HR).toBe(10);
    expect(byCat(agg('max')).HR).toBe(30);
    expect(byCat(agg('median')).HR).toBe(20);
    expect(byCat(agg('countDistinct')).HR).toBe(2);
  });

  it('names the series after the measure when there is no series column', () => {
    expect(aggregate(ds, ALL, spec()).series[0].name).toBe('Amount');
  });
});

// Ruling F3 -- the tile spec carries `encoding.x.bin` and the suggestion
// engine sets it, but nothing wired it THROUGH the aggregator. Untested
// wiring between two tested units: without it a time series groups by
// exact timestamp and produces one category per row.
describe('aggregate with encoding.x.bin', () => {
  const dated = {
    headers: ['Raised', 'Amount'],
    rows: [
      ['01/01/2024', '10'], ['15/01/2024', '20'],
      ['03/02/2024', '30'], ['20/02/2024', '40'],
    ],
  };
  // A real plan, not an empty one: date strings only become epochs when
  // the castType step runs, so an empty plan would leave every value
  // unparsed and the test would pass or fail for the wrong reason.
  const datedProfile = profileDataset(dated);
  const dsDated = applyCleanPlan(dated, proposeCleanPlan(datedProfile, dated), datedProfile);

  it('truncates a temporal x column to the requested unit', () => {
    const r = aggregate(dsDated, new Uint8Array(4).fill(1), {
      chart: 'line',
      encoding: {
        x: { column: 'Raised', bin: 'month' },
        y: [{ column: 'Amount', agg: 'sum' }],
        series: null,
      },
      sort: { by: 'x', dir: 'asc' },
      limit: 20,
    });
    expect(r.categories).toHaveLength(2);
    expect(r.series[0].data).toEqual([30, 70]);
  });

  it('groups by day when asked, not by exact timestamp', () => {
    const r = aggregate(dsDated, new Uint8Array(4).fill(1), {
      chart: 'line',
      encoding: {
        x: { column: 'Raised', bin: 'day' },
        y: [{ column: 'Amount', agg: 'count' }],
        series: null,
      },
      sort: { by: 'x', dir: 'asc' },
      limit: 20,
    });
    expect(r.categories).toHaveLength(4);
  });

  it('bins a numeric x column instead of making one category per value', () => {
    const spread = {
      headers: ['Score', 'Amount'],
      // Deliberately NOT 0..39: a consecutive integer run is an
      // identifier by spec 7.3, and identifiers are never measures.
      rows: Array.from({ length: 40 }, (_, i) => [String(i * 2.5), '1']),
    };
    const spreadProfile = profileDataset(spread);
    const dsSpread = applyCleanPlan(spread, [], spreadProfile);
    const r = aggregate(dsSpread, new Uint8Array(40).fill(1), {
      chart: 'histogram',
      encoding: {
        x: { column: 'Score', bin: 'auto' },
        y: [{ column: 'Amount', agg: 'count' }],
        series: null,
      },
      sort: { by: 'x', dir: 'asc' },
      limit: 50,
    });
    expect(r.categories.length).toBeGreaterThan(1);
    expect(r.categories.length).toBeLessThan(40);
    // Every row lands in exactly one bin, so the counts must still add up.
    expect(r.series[0].data.reduce((a, b) => a + b, 0)).toBe(40);
  });
});

describe('truncateDate', () => {
  const t = Date.UTC(2024, 4, 17, 13, 45); // 17 May 2024
  it('truncates to day', () => expect(truncateDate(t, 'day')).toBe(Date.UTC(2024, 4, 17)));
  it('truncates to month', () => expect(truncateDate(t, 'month')).toBe(Date.UTC(2024, 4, 1)));
  it('truncates to quarter', () => expect(truncateDate(t, 'quarter')).toBe(Date.UTC(2024, 3, 1)));
  it('truncates to year', () => expect(truncateDate(t, 'year')).toBe(Date.UTC(2024, 0, 1)));
});

describe('chooseTruncation', () => {
  const DAY = 86400000;
  it('uses day below 90 days', () => {
    expect(chooseTruncation(0, 60 * DAY)).toBe('day');
  });
  it('uses month below 3 years', () => {
    expect(chooseTruncation(0, 500 * DAY)).toBe('month');
  });
  it('uses quarter beyond 3 years', () => {
    expect(chooseTruncation(0, 2000 * DAY)).toBe('quarter');
  });
});

describe('binNumeric', () => {
  it('covers the whole range with edges and matching labels', () => {
    const values = Float64Array.from({ length: 50 }, (_, i) => i);
    const { edges, labels } = binNumeric(values, new Uint8Array(50).fill(1));
    expect(edges.length).toBe(labels.length + 1);
    expect(edges[0]).toBeLessThanOrEqual(0);
    expect(edges[edges.length - 1]).toBeGreaterThanOrEqual(49);
  });

  it('falls back to a single bin when every value is identical', () => {
    const values = Float64Array.from({ length: 10 }, () => 7);
    expect(binNumeric(values, new Uint8Array(10).fill(1)).labels).toHaveLength(1);
  });

  it('ignores masked-out and missing values', () => {
    const values = Float64Array.from([1, NaN, 100]);
    const { edges } = binNumeric(values, new Uint8Array([1, 1, 0]));
    expect(edges[edges.length - 1]).toBeLessThan(100);
  });
});

describe('aggregate over a multi column', () => {
  const dataset = buildDataset({
    headers: ['Challenges'],
    columns: [['A;B;', 'B;', 'A;B;C;', '']],
    profile: { columns: [{ name: 'Challenges', type: 'multi', role: 'dimension', separator: ';' }] },
  });

  it('counts each option a row picked', () => {
    const result = aggregate(dataset, null, {
      encoding: { x: { column: 'Challenges' }, y: [{ column: null, agg: 'count' }] },
      sort: { by: 'y', dir: 'desc' },
    });
    const counts = Object.fromEntries(
      result.categories.map((c, i) => [c, result.series[0].data[i]]),
    );
    expect(counts).toEqual({ B: 3, A: 2, C: 1 });
  });
});
