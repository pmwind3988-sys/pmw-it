import { describe, it, expect } from 'vitest';
import { analysisTiles, withAnalysisTiles, isAnalysisTile } from './analysisTiles.js';
import { DERIVED_HEADERS } from './deriveColumns.js';
import { validateTileSpec } from '../canvas/chartSpecs.js';
import { buildDataset } from '../engine/dataset.js';

const dataset = buildDataset({
  headers: [...DERIVED_HEADERS, 'Department'],
  columns: [
    ['Access', 'Hardware'],
    ['Access', 'Hardware;Access'],
    ['Slow laptops', 'Slow laptops'],
    [1, 2],
    [40, 80],
    ['HR', 'IT'],
  ],
  profile: {
    columns: [
      { name: 'Issue category', type: 'categorical', role: 'dimension' },
      { name: 'Issue categories', type: 'multi', role: 'dimension' },
      { name: 'Theme', type: 'categorical', role: 'dimension' },
      { name: 'Issue count', type: 'numeric', role: 'measure' },
      { name: 'Severity', type: 'numeric', role: 'measure' },
      { name: 'Department', type: 'categorical', role: 'dimension' },
    ],
  },
});

describe('analysisTiles', () => {
  it('builds a dashboard every tile of which this dataset can draw', () => {
    const tiles = analysisTiles();
    expect(tiles.length).toBeGreaterThan(3);
    for (const tile of tiles) {
      expect(validateTileSpec(tile, dataset)).toMatchObject({ ok: true });
    }
  });

  it('leads with single-number cards, then the charts', () => {
    const kinds = analysisTiles().map((t) => t.chart);
    expect(kinds.slice(0, 3)).toEqual(['kpi', 'kpi', 'kpi']);
    expect(kinds.slice(3).every((k) => k === 'bar')).toBe(true);
  });

  it('counts categories on the multi column, so one person counts once per category', () => {
    const categories = analysisTiles().find((t) => t.id === 'txt_categories');
    expect(categories.encoding.x.column).toBe('Issue categories');
  });

  it('averages severity rather than summing it', () => {
    // Summed, a category twenty people mentioned mildly would outrank
    // one three people are furious about.
    const severity = analysisTiles().find((t) => t.id === 'txt_severity');
    expect(severity.encoding.y[0]).toMatchObject({ column: 'Severity', agg: 'avg' });
  });

  it('drops the tile rather than the column when a column is missing', () => {
    const tiles = analysisTiles(['Theme']);
    expect(tiles.map((t) => t.id)).toEqual(['txt_kpi_people', 'txt_themes']);
  });

  it('gives every tile a stable id', () => {
    expect(analysisTiles().map((t) => t.id)).toEqual(analysisTiles().map((t) => t.id));
    expect(analysisTiles().every(isAnalysisTile)).toBe(true);
  });
});

describe('withAnalysisTiles', () => {
  const mine = [{ id: 'sug_1', title: 'Mine' }];

  it('keeps charts the user already had, and leads with the analysis', () => {
    const merged = withAnalysisTiles(mine);
    expect(merged[merged.length - 1]).toEqual(mine[0]);
    expect(isAnalysisTile(merged[0])).toBe(true);
  });

  it('replaces its own tiles on a second run instead of stacking them', () => {
    const once = withAnalysisTiles(mine);
    const twice = withAnalysisTiles(once);
    expect(twice.map((t) => t.id)).toEqual(once.map((t) => t.id));
  });
});
