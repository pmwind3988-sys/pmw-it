// Tile spec + aggregate result -> an ECharts option -- spec §10.1, §10.2.
//
// Pure functions, no React and no ECharts import. That is what makes
// every mapping decision in here testable without a DOM: the tests
// assert on the option object, not on pixels.
//
// Two of the eight chart types deliberately do NOT return an ECharts
// option. A KPI is a number and a table is rows; wrapping either in a
// chart config to make the return type uniform would mean building a
// chart in order to render text.

// What each chart type needs from a tile spec before it can be drawn.
// `y` is a COUNT of measures, so scatter can say "two" and mean it.
export const CHART_TYPES = [
  { id: 'bar', label: 'Bar', needs: { x: 1, y: 1, series: 0 } },
  { id: 'line', label: 'Line', needs: { x: 1, y: 1, series: 0 } },
  { id: 'area', label: 'Area', needs: { x: 1, y: 1, series: 0 } },
  { id: 'pie', label: 'Pie', needs: { x: 1, y: 1, series: 0 } },
  { id: 'scatter', label: 'Scatter', needs: { x: 1, y: 2, series: 0 } },
  { id: 'heatmap', label: 'Heatmap', needs: { x: 1, y: 1, series: 1 } },
  { id: 'kpi', label: 'Single number', needs: { x: 0, y: 1, series: 0 } },
  { id: 'table', label: 'Table', needs: { x: 1, y: 1, series: 0 } },
];

// One stack name shared by every series in a stacked tile. ECharts
// stacks series that share this string, so the value is arbitrary but
// must be identical across them.
const STACK_GROUP = 'total';

export function chartTypeById(id) {
  return CHART_TYPES.find((c) => c.id === id);
}

/**
 * Whether a tile can be drawn against this dataset (spec §12).
 *
 * The failure case that matters is a chart whose dataset has
 * changed underneath it. `reason` names the missing column, because
 * "This tile is invalid" leaves the user with nothing to act on, while
 * "Column 'Cost centre' is not in this dataset" tells them exactly what
 * to re-import or re-map.
 */
export function validateTileSpec(spec, dataset) {
  if (!dataset?.byName) return { ok: false, reason: 'There is no dataset loaded.' };

  const encoding = spec?.encoding ?? {};
  const wanted = [
    encoding.x?.column,
    ...(encoding.y ?? []).map((m) => m.column),
    encoding.series?.column,
  ].filter(Boolean);

  for (const name of wanted) {
    if (!dataset.byName.has(name)) {
      return { ok: false, reason: `Column "${name}" is not in this dataset.` };
    }
  }

  const type = chartTypeById(spec?.chart);
  if (!type) return { ok: false, reason: `Unknown chart type "${spec?.chart}".` };

  const measures = (encoding.y ?? []).length;
  if (measures < type.needs.y) {
    return {
      ok: false,
      reason: `A ${type.label.toLowerCase()} chart needs ${type.needs.y} measures; this tile has ${measures}.`,
    };
  }
  if (type.needs.x > 0 && !encoding.x?.column) {
    return { ok: false, reason: `A ${type.label.toLowerCase()} chart needs a column on its X axis.` };
  }

  return { ok: true };
}

function axisLabel(spec) {
  return spec?.encoding?.x?.column ?? '';
}

function measureLabel(spec) {
  const measure = spec?.encoding?.y?.[0];
  return measure?.column ?? 'Count';
}

function baseGrid() {
  // Tight, because a tile is small. `containLabel` lets ECharts reserve
  // exactly the room the tick labels need rather than a fixed guess,
  // which is what stops long category names being clipped.
  return {
    left: 8, right: 12, top: 8, bottom: 8, containLabel: true,
  };
}

// Dims the marks that are not part of the current selection.
//
// This is presentation ONLY (spec §10.3). The source tile of a
// cross-filter keeps every category -- that is the whole rule -- so the
// numbers are untouched and only the opacity says which one was
// clicked. Recomputing the tile against its own selection is exactly
// the bug the self-exclusion rule exists to prevent.
const DIMMED_OPACITY = 0.28;

function withSelection(data, categories, selected) {
  if (!selected || selected.length === 0) return data;
  const wanted = new Set(selected);
  return data.map((value, i) => (wanted.has(categories[i])
    ? { value, itemStyle: { opacity: 1 } }
    : { value, itemStyle: { opacity: DIMMED_OPACITY } }));
}

function cartesian(type, aggResult, spec, selected) {
  const stacked = Boolean(spec?.stacked);
  const isArea = type === 'area';

  return {
    grid: baseGrid(),
    tooltip: { trigger: 'axis' },
    legend: { data: aggResult.series.map((s) => s.name), type: 'scroll' },
    xAxis: { type: 'category', data: aggResult.categories, name: axisLabel(spec) },
    yAxis: { type: 'value' },
    series: aggResult.series.map((s) => ({
      name: s.name,
      // 'area' is not an ECharts series type -- an area chart is a line
      // that fills, so the type stays 'line' and areaStyle does the work.
      type: isArea ? 'line' : type,
      data: withSelection(s.data, aggResult.categories, selected),
      ...(isArea ? { areaStyle: {}, smooth: false } : {}),
      ...(stacked ? { stack: STACK_GROUP } : {}),
      // Selection state is what lets the SOURCE tile of a cross-filter
      // highlight the bar that was clicked while keeping the rest
      // visible -- the visual half of the self-exclusion rule.
      selectedMode: 'multiple',
      emphasis: { focus: 'series' },
      select: { itemStyle: { borderWidth: 2 } },
    })),
  };
}

function pie(aggResult, spec, selected) {
  const first = aggResult.series[0] ?? { data: [] };
  const wanted = new Set(selected ?? []);
  return {
    tooltip: { trigger: 'item' },
    legend: { data: aggResult.categories, type: 'scroll' },
    series: [{
      name: measureLabel(spec),
      type: 'pie',
      radius: ['45%', '72%'],
      // A donut, not a filled pie: the hole gives the labels somewhere
      // to sit and makes small slices easier to compare by arc length.
      avoidLabelOverlap: true,
      label: { show: false },
      data: aggResult.categories.map((name, i) => ({
        name,
        value: first.data[i] ?? 0,
        ...(wanted.size > 0
          ? { itemStyle: { opacity: wanted.has(name) ? 1 : DIMMED_OPACITY } }
          : null),
      })),
      selectedMode: 'multiple',
      emphasis: { focus: 'self' },
    }],
  };
}

function scatter(aggResult, spec) {
  const [xs, ys] = aggResult.series;
  return {
    grid: baseGrid(),
    tooltip: { trigger: 'item' },
    legend: { data: aggResult.series.map((s) => s.name), type: 'scroll' },
    xAxis: { type: 'value', name: xs?.name ?? axisLabel(spec) },
    yAxis: { type: 'value', name: ys?.name ?? '' },
    series: [{
      name: ys?.name ?? measureLabel(spec),
      type: 'scatter',
      data: (xs?.data ?? []).map((x, i) => [x, ys?.data?.[i] ?? 0]),
      selectedMode: 'multiple',
      emphasis: { focus: 'series' },
    }],
  };
}

function heatmap(aggResult, spec) {
  const data = [];
  let max = 0;
  aggResult.series.forEach((s, y) => {
    s.data.forEach((value, x) => {
      data.push([x, y, value]);
      if (value > max) max = value;
    });
  });

  return {
    grid: baseGrid(),
    tooltip: { trigger: 'item' },
    xAxis: { type: 'category', data: aggResult.categories, name: axisLabel(spec) },
    yAxis: { type: 'category', data: aggResult.series.map((s) => s.name) },
    visualMap: {
      min: 0, max: max || 1, calculable: true, orient: 'horizontal', left: 'center', bottom: 0,
    },
    series: [{
      name: measureLabel(spec),
      type: 'heatmap',
      data,
      selectedMode: 'multiple',
      emphasis: { itemStyle: { borderWidth: 2 } },
    }],
  };
}

/**
 * Maps an aggregate result onto whatever the chart type needs.
 *
 * KPI and table return their own shapes rather than ECharts options --
 * see the note at the top of the file.
 */
export function toEChartsOption(chartType, aggResult, spec, selected = null) {
  const result = {
    categories: aggResult?.categories ?? [],
    series: aggResult?.series ?? [],
  };

  switch (chartType) {
    case 'kpi': {
      const first = result.series[0] ?? { data: [] };
      return {
        kind: 'kpi',
        // The total across every category, which is what a KPI on a
        // grouped aggregate means. An empty result is 0, not NaN -- a
        // tile reading "NaN" tells the user nothing is wrong with their
        // data when nothing is.
        value: first.data.reduce((a, b) => a + (Number.isFinite(b) ? b : 0), 0),
        label: measureLabel(spec),
      };
    }
    case 'table':
      return {
        kind: 'table',
        headers: [axisLabel(spec), ...result.series.map((s) => s.name)],
        rows: result.categories.map((category, i) => [
          category, ...result.series.map((s) => s.data[i] ?? 0),
        ]),
      };
    case 'pie':
      return pie(result, spec, selected);
    case 'scatter':
      return scatter(result, spec);
    case 'heatmap':
      return heatmap(result, spec);
    default:
      return cartesian(chartType, result, spec, selected);
  }
}
