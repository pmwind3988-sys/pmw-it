import { describe, it, expect } from 'vitest';
import { suggestCharts } from './suggestCharts.js';
import { profileDataset } from '../profile/profileDataset.js';

const grid = {
  headers: ['Dept', 'Created', 'Amount'],
  rows: Array.from({ length: 40 }, (_, i) => [
    ['HR', 'IT', 'Finance'][i % 3],
    `${String((i % 28) + 1).padStart(2, '0')}/01/2024`,
    String((i + 1) * 10),
  ]),
};

describe('suggestCharts', () => {
  const tiles = suggestCharts(profileDataset(grid));

  it('returns at most six tiles', () => {
    expect(tiles.length).toBeGreaterThan(0);
    expect(tiles.length).toBeLessThanOrEqual(6);
  });

  it('leads with a KPI row', () => {
    expect(tiles[0].chart).toBe('kpi');
  });

  it('includes a time series on the primary temporal column', () => {
    const line = tiles.find((t) => t.chart === 'line');
    expect(line.encoding.x.column).toBe('Created');
  });

  it('truncates a sub-90-day span to day', () => {
    expect(tiles.find((t) => t.chart === 'line').encoding.x.bin).toBe('day');
  });

  it('includes a bar chart for the low-cardinality dimension', () => {
    const bar = tiles.find((t) => t.chart === 'bar' && t.encoding.x.column === 'Dept');
    expect(bar.encoding.y[0].column).toBe('Amount');
  });

  it('gives every tile a unique id and a title', () => {
    expect(new Set(tiles.map((t) => t.id)).size).toBe(tiles.length);
    expect(tiles.every((t) => t.title)).toBe(true);
  });

  it('never suggests an identifier as a measure', () => {
    const withId = {
      headers: ['Employee ID', 'Dept'],
      rows: Array.from({ length: 30 }, (_, i) => [String(1000 + i), 'HR']),
    };
    const t = suggestCharts(profileDataset(withId));
    expect(t.every((x) => (x.encoding.y ?? []).every((y) => y.column !== 'Employee ID')))
      .toBe(true);
  });

  it('returns an empty array when nothing is chartable', () => {
    const junk = {
      headers: ['Notes'],
      rows: Array.from({ length: 50 }, (_, i) => [`free text ${i}`]),
    };
    expect(suggestCharts(profileDataset(junk))).toEqual([]);
  });

  // Every suggested tile is handed straight to the canvas, so a spec the
  // validator rejects would land the user on a dashboard of error
  // messages.
  it('produces tiles that all pass validation', async () => {
    const { validateTileSpec } = await import('../canvas/chartSpecs.js');
    const dataset = { byName: new Map(grid.headers.map((h) => [h, {}])) };
    for (const tile of tiles) {
      expect(validateTileSpec(tile, dataset)).toMatchObject({ ok: true });
    }
  });

  it('sorts a bar chart by value and caps it, so a wide dimension stays readable', () => {
    const bar = tiles.find((t) => t.chart === 'bar' && t.encoding.x.column === 'Dept');
    expect(bar.sort).toMatchObject({ by: 'y', dir: 'desc' });
    expect(bar.limit).toBeLessThanOrEqual(10);
  });

  // A time series sorted by value is not a time series.
  it('sorts the time series chronologically, not by value', () => {
    expect(tiles.find((t) => t.chart === 'line').sort).toMatchObject({ by: 'x', dir: 'asc' });
  });

  // Requires the profile to carry min/max for TEMPORAL columns, not just
  // numeric ones. Without those, chooseTruncation has nothing to choose
  // from and every time series falls back to a day grain -- which on
  // five years of data draws about eighteen hundred categories.
  it('picks a month grain for a span of a couple of years', () => {
    const spread = {
      headers: ['Created', 'Amount'],
      rows: Array.from({ length: 40 }, (_, i) => [
        `15/${String((i % 12) + 1).padStart(2, '0')}/${2022 + Math.floor(i / 24)}`,
        String((i + 1) * 10),
      ]),
    };
    const line = suggestCharts(profileDataset(spread)).find((t) => t.chart === 'line');
    expect(line.encoding.x.bin).toBe('month');
  });

  it('picks a quarter grain once the span passes three years', () => {
    const long = {
      headers: ['Created', 'Amount'],
      rows: Array.from({ length: 48 }, (_, i) => [
        `15/${String((i % 12) + 1).padStart(2, '0')}/${2020 + Math.floor(i / 12)}`,
        String((i + 1) * 10),
      ]),
    };
    const line = suggestCharts(profileDataset(long)).find((t) => t.chart === 'line');
    expect(line.encoding.x.bin).toBe('quarter');
  });

  // Spec §10.5 -- a dimension with hundreds of distinct values makes a
  // bar chart of hundreds of bars, which is a smear.
  it('does not suggest a bar chart for a high-cardinality dimension', () => {
    const wide = {
      headers: ['Ticket', 'Amount'],
      rows: Array.from({ length: 60 }, (_, i) => [`T-${i}`, String(i * 3 + 1)]),
    };
    const t = suggestCharts(profileDataset(wide));
    expect(t.some((x) => x.chart === 'bar' && x.encoding.x.column === 'Ticket')).toBe(false);
  });

  it('gives the row-count KPI first place, before the measure KPIs', () => {
    expect(tiles[0]).toMatchObject({ chart: 'kpi' });
    expect(tiles[0].title.toLowerCase()).toContain('row');
  });
});

// The file-title nudge. Two dimensions of comparable shape, one of which
// the title names -- without the nudge the order is decided by fill and
// cardinality alone, which is how a dashboard about departments leads
// with a chart about something else.
describe('suggestCharts with a file-title focus', () => {
  const grid = {
    headers: ['Region', 'Department', 'Hours'],
    rows: [
      ['North', 'Finance', '5'], ['South', 'Finance', '3'], ['East', 'Logistics', '8'],
      ['West', 'Logistics', '2'], ['North', 'Sales', '6'], ['South', 'Sales', '4'],
      ['East', 'QAQC', '7'], ['West', 'QAQC', '1'],
    ],
  };
  const profile = profileDataset(grid);
  const barFor = (focus) => suggestCharts(profile, null, focus)
    .filter((t) => t.chart === 'bar' && t.encoding.x?.column)
    .map((t) => t.encoding.x.column);

  it('puts the dimension the title names in front', () => {
    expect(barFor(['department'])[0]).toBe('Department');
  });

  it('changes nothing when no focus is given', () => {
    expect(barFor([])).toEqual(barFor(undefined));
  });

  it('never drops a chart the focus does not mention', () => {
    expect(barFor(['department'])).toContain('Region');
  });
});
