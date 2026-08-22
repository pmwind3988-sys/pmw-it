import { describe, it, expect } from 'vitest';
import { toEChartsOption, validateTileSpec, CHART_TYPES } from './chartSpecs.js';

const agg = {
  categories: ['HR', 'IT'],
  series: [{ name: 'Amount', data: [40, 60] }],
};
const tile = {
  id: 't1',
  title: 'By dept',
  chart: 'bar',
  encoding: { x: { column: 'Dept' }, y: [{ column: 'Amount', agg: 'sum' }], series: null },
};

describe('toEChartsOption', () => {
  it('maps categories onto the x axis for a bar chart', () => {
    const o = toEChartsOption('bar', agg, tile);
    expect(o.xAxis.data).toEqual(['HR', 'IT']);
    expect(o.series[0].type).toBe('bar');
    expect(o.series[0].data).toEqual([40, 60]);
  });

  it('emits a line series for line charts', () => {
    expect(toEChartsOption('line', agg, tile).series[0].type).toBe('line');
  });

  it('fills the area for area charts', () => {
    expect(toEChartsOption('area', agg, tile).series[0].areaStyle).toBeDefined();
  });

  it('reshapes categories and values into name/value pairs for pie', () => {
    const o = toEChartsOption('pie', agg, tile);
    expect(o.series[0].data).toEqual([
      { name: 'HR', value: 40 }, { name: 'IT', value: 60 },
    ]);
  });

  it('enables select and blur states so the source tile can highlight', () => {
    const o = toEChartsOption('bar', agg, tile);
    expect(o.series[0].selectedMode).toBeTruthy();
  });

  it('stacks series when the tile asks for it', () => {
    const stacked = { ...tile, stacked: true };
    const multi = {
      categories: ['HR'],
      series: [{ name: 'a', data: [1] }, { name: 'b', data: [2] }],
    };
    const o = toEChartsOption('bar', multi, stacked);
    expect(o.series[0].stack).toBe(o.series[1].stack);
  });

  it('reduces a kpi tile to a single number', () => {
    const o = toEChartsOption('kpi', agg, { ...tile, chart: 'kpi' });
    expect(o.value).toBe(100);
  });

  // An area chart is a line chart that fills, so it must still BE a
  // line to ECharts -- 'area' is not a series type.
  it('emits area charts as line series, since ECharts has no area type', () => {
    expect(toEChartsOption('area', agg, tile).series[0].type).toBe('line');
  });

  it('does not stack series unless asked', () => {
    const multi = {
      categories: ['HR'],
      series: [{ name: 'a', data: [1] }, { name: 'b', data: [2] }],
    };
    expect(toEChartsOption('bar', multi, tile).series[0].stack).toBeUndefined();
  });

  it('emits one legend entry per series', () => {
    const multi = {
      categories: ['HR'],
      series: [{ name: 'a', data: [1] }, { name: 'b', data: [2] }],
    };
    expect(toEChartsOption('bar', multi, tile).legend.data).toEqual(['a', 'b']);
  });

  it('gives a table tile its rows rather than an ECharts option', () => {
    const o = toEChartsOption('table', agg, { ...tile, chart: 'table' });
    expect(o.rows).toEqual([['HR', 40], ['IT', 60]]);
    expect(o.headers).toEqual(['Dept', 'Amount']);
  });

  it('survives an empty aggregate rather than throwing', () => {
    const empty = { categories: [], series: [{ name: 'Amount', data: [] }] };
    expect(() => toEChartsOption('bar', empty, tile)).not.toThrow();
    expect(toEChartsOption('kpi', empty, { ...tile, chart: 'kpi' }).value).toBe(0);
  });
});

describe('validateTileSpec', () => {
  const dataset = { byName: new Map([['Dept', {}], ['Amount', {}]]) };

  it('accepts a tile whose columns all exist', () => {
    expect(validateTileSpec(tile, dataset).ok).toBe(true);
  });

  // Spec §12 -- a tile pointing at a deleted column must explain itself.
  it('rejects a tile referencing a missing column and names it', () => {
    const broken = { ...tile, encoding: { ...tile.encoding, x: { column: 'Gone' } } };
    const r = validateTileSpec(broken, dataset);
    expect(r.ok).toBe(false);
    expect(r.reason).toContain('Gone');
  });

  it('names a missing measure column too, not only a missing x', () => {
    const broken = { ...tile, encoding: { ...tile.encoding, y: [{ column: 'Ghost', agg: 'sum' }] } };
    const r = validateTileSpec(broken, dataset);
    expect(r.ok).toBe(false);
    expect(r.reason).toContain('Ghost');
  });

  it('names a missing series column', () => {
    const broken = { ...tile, encoding: { ...tile.encoding, series: { column: 'Nope' } } };
    expect(validateTileSpec(broken, dataset).reason).toContain('Nope');
  });

  it('rejects a chart type that needs two measures but has one', () => {
    const scatter = { ...tile, chart: 'scatter' };
    expect(validateTileSpec(scatter, dataset).ok).toBe(false);
  });

  it('rejects a tile with no dataset at all rather than throwing', () => {
    expect(validateTileSpec(tile, null).ok).toBe(false);
  });
});

describe('CHART_TYPES', () => {
  it('declares what each chart type needs', () => {
    expect(CHART_TYPES.find((c) => c.id === 'scatter').needs.y).toBe(2);
  });

  it('covers every chart type the plan names', () => {
    expect(CHART_TYPES.map((c) => c.id).sort()).toEqual(
      ['area', 'bar', 'heatmap', 'kpi', 'line', 'pie', 'scatter', 'table'],
    );
  });
});
