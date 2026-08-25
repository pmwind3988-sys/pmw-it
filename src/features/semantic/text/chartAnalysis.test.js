import { describe, it, expect } from 'vitest';
import { withAnalysisCharted } from './chartAnalysis.js';
import { DERIVED_HEADERS } from './deriveColumns.js';
import { isAnalysisTile } from './analysisTiles.js';

const analysis = {
  columnName: 'Describe the biggest issue you face',
  buckets: [
    { id: 'sap', label: 'SAP / ERP' },
    { id: 'network', label: 'Network & Internet' },
  ],
  themes: [{ id: 't1', name: 'slow · posting' }],
  fragments: [
    { id: '0:0', row: 0, text: 'sap is slow', severity: 0.8, bucketId: 'sap', themeId: 't1', noise: false },
    { id: '0:1', row: 0, text: 'wifi drops', severity: 0.4, bucketId: 'network', themeId: null, noise: false },
    { id: '1:0', row: 1, text: 'wifi drops', severity: 0.5, bucketId: 'network', themeId: null, noise: false },
  ],
};

const base = {
  analysis,
  grid: {
    headers: ['Email', 'Which department are you in?', 'Describe the biggest issue you face'],
    rows: [
      ['ali@pmw-group.com', 'Finance', 'SAP is slow. The wifi drops.'],
      ['bella@pmw-group.com', 'HR', 'The wifi drops.'],
      ['chan@pmw-group.com', 'HR', ''],
    ],
  },
  tiles: [{ id: 'tile_1', title: 'Responses by department', chart: 'bar' }],
};

describe('withAnalysisCharted', () => {
  const next = withAnalysisCharted(base);

  it('adds the reading to the sheet as ordinary columns', () => {
    expect(next.grid.headers.slice(-DERIVED_HEADERS.length)).toEqual(DERIVED_HEADERS);
    expect(next.grid.rows).toHaveLength(3);
    for (const row of next.grid.rows) expect(row).toHaveLength(8);
  });

  it('keeps every column the form already had', () => {
    expect(next.grid.headers.slice(0, 3)).toEqual(base.grid.headers);
    expect(next.grid.rows[0][0]).toBe('ali@pmw-group.com');
  });

  it('charts the categories, and keeps the charts already on screen', () => {
    expect(next.tiles.some(isAnalysisTile)).toBe(true);
    expect(next.tiles.map((t) => t.id)).toContain('tile_1');
  });

  it('re-profiles rather than patching, so the new columns are typed', () => {
    const categories = next.profile.columns.find((c) => c.name === 'Issue categories');
    expect(categories.type).toBe('multi');
  });

  it('replaces the analysis rather than stacking a second copy of it', () => {
    const twice = withAnalysisCharted(next);
    expect(twice.grid.headers).toEqual(next.grid.headers);
    expect(twice.tiles.filter(isAnalysisTile)).toHaveLength(
      next.tiles.filter(isAnalysisTile).length,
    );
    expect(twice.tiles.map((t) => t.id)).toContain('tile_1');
  });

  it('latches, so the caller can tell a first reading from a re-score', () => {
    expect(base.autoCharted).toBeUndefined();
    expect(next.autoCharted).toBe(true);
  });

  it('leaves the state alone when there is nothing to chart', () => {
    expect(withAnalysisCharted({ grid: base.grid })).toEqual({ grid: base.grid });
    expect(withAnalysisCharted(null)).toBe(null);
  });
});
