// Profile -> a ranked set of starter tiles -- spec §10.5.
//
// The user lands on a populated canvas, not an empty "Add chart" page.
// That is the difference between a tool that shows you your data and a
// tool that asks you to describe it first.
//
// `topMeasure` and `primaryTemporal` come from the profile and are NOT
// re-derived here. Two places deciding independently which measure is
// "the" measure is how a KPI ends up disagreeing with the chart beside
// it.

import { chooseTruncation } from '../engine/aggregate.js';
import { relevanceOf } from '../intent/fileIntent.js';

// How much a column named in the file title can outrank one that is
// merely well shaped. Capped low deliberately: a title is a hint about
// what the reader cares about, not evidence about the data, and a
// two-value column should not take the first slot just because the file
// name happens to repeat its header.
const FOCUS_BONUS = 0.5;

export const MAX_TILES = 6;

// Spec §10.5's readable band for a bar chart's x axis. Below 3 there is
// barely a comparison to make; above 12 the labels start colliding and
// the chart becomes a smear.
const CARDINALITY_MIN = 3;
const CARDINALITY_MAX = 12;

// Above this a dimension is not a category any more, whatever its type
// says -- a bar per ticket number is not a chart.
const BAR_CARDINALITY_LIMIT = 12;

const MAX_MEASURE_KPIS = 3;

// Correlation strong enough to be worth a scatter tile. Spec §10.5.
const SCATTER_MIN_R = 0.3;

// Correlation is a ranking heuristic here, not a reported statistic, so
// a sample is enough and a full pass over 100k rows twice per measure
// pair is not worth it.
const CORRELATION_SAMPLE = 5000;

function tile(id, title, chart, encoding, extra = {}) {
  return {
    id,
    title,
    chart,
    encoding: { x: null, y: [], series: null, ...encoding },
    sort: { by: 'y', dir: 'desc' },
    limit: 10,
    size: 'M',
    stacked: false,
    // Tiles respond to cross-filters by default. Nothing turns this
    // off any more -- the flag is still read, because a chart that
    // should show an unfiltered total is a real thing to want and the
    // machinery for it costs one property.
    respondsToFilters: true,
    ...extra,
  };
}

// How interesting a dimension is: peaks inside the readable cardinality
// band and falls off outside it, scaled by how full the column is. A
// column that is 90% empty makes a chart that is 90% nothing.
function dimensionScore(column) {
  const n = column.distinctCount;
  let shape;
  if (n < CARDINALITY_MIN) {
    // Not zero: a two-category split is still a real comparison, just a
    // thin one.
    shape = 0.4 * (n / CARDINALITY_MIN);
  } else if (n <= CARDINALITY_MAX) {
    shape = 1;
  } else {
    shape = CARDINALITY_MAX / n;
  }
  return shape * column.nonNullRatio;
}

// A measure whose values barely vary draws a flat line, so the
// coefficient of variation stands in for "is there anything to see".
function measureScore(column) {
  const mean = column.mean;
  const spread = (column.max ?? 0) - (column.min ?? 0);
  if (!Number.isFinite(mean) || mean === 0) {
    return column.nonNullRatio * (spread > 0 ? 0.5 : 0);
  }
  const cv = Math.abs(spread / mean);
  return column.nonNullRatio * Math.min(1, 0.3 + cv);
}

// Pearson r over a sample. Returns 0 when either column is constant,
// since r is undefined there and "no relationship worth charting" is the
// right answer either way.
function correlation(a, b, rowCount) {
  const step = Math.max(1, Math.floor(rowCount / CORRELATION_SAMPLE));
  let n = 0;
  let sa = 0;
  let sb = 0;
  let saa = 0;
  let sbb = 0;
  let sab = 0;

  for (let i = 0; i < rowCount; i += step) {
    const x = a[i];
    const y = b[i];
    if (Number.isNaN(x) || Number.isNaN(y)) continue;
    n++;
    sa += x;
    sb += y;
    saa += x * x;
    sbb += y * y;
    sab += x * y;
  }

  if (n < 3) return 0;
  const cov = sab / n - (sa / n) * (sb / n);
  const va = saa / n - (sa / n) ** 2;
  const vb = sbb / n - (sb / n) ** 2;
  if (va <= 0 || vb <= 0) return 0;
  return cov / Math.sqrt(va * vb);
}

/**
 * Ranked starter tiles for a freshly profiled dataset.
 *
 * `dataset` is optional. Scatter suggestions need the actual values to
 * compute a correlation, and the profile does not carry them; without a
 * dataset those candidates are simply skipped rather than guessed at.
 *
 * `focus` is optional too -- the keywords read out of the file name by
 * `intent/fileIntent.js`. It only ever nudges a candidate up the order;
 * with no focus the ranking is exactly what it was before.
 *
 * `written` is the names of the columns holding written answers, from
 * `text/detectTextColumns.js`. They are excluded outright: what those
 * answers say is the analysis's job, and a bar chart of them is a bar
 * per sentence. Cardinality alone does not catch this -- on a sheet of
 * twelve responses twelve paragraphs sit comfortably under the limit --
 * and neither does the column TYPE, since a column of a dozen distinct
 * sentences profiles as categorical.
 */
export function suggestCharts(profile, dataset = null, focus = [], written = []) {
  const columns = profile?.columns ?? [];
  const measures = columns.filter((c) => c.role === 'measure');
  const dimensions = columns.filter((c) => c.role === 'dimension');
  const temporals = columns.filter((c) => c.role === 'temporal');

  // A dimension only earns a chart if it is narrow enough to draw. A
  // column of 200 distinct ticket numbers is categorical by type and
  // useless as an axis.
  const prose = new Set(written);
  const chartableDimensions = dimensions.filter(
    (c) => !prose.has(c.name) && c.distinctCount <= BAR_CARDINALITY_LIMIT,
  );

  // Nothing here would produce an actual chart -- a sheet of free text,
  // say. An empty canvas with an "add a chart" prompt is honest; a lone
  // row-count tile dressed up as a dashboard is not.
  if (measures.length === 0 && temporals.length === 0 && chartableDimensions.length === 0) {
    return [];
  }

  const topMeasure = columns.find((c) => c.name === profile.topMeasure) ?? null;
  const primaryTemporal = columns.find((c) => c.name === profile.primaryTemporal) ?? null;

  let seq = 0;
  const nextId = () => {
    seq += 1;
    return `sug_${seq}`;
  };

  // --- the KPI row, always first -------------------------------------
  //
  // Not scored: it is the summary line of the dashboard, and a summary
  // that sorts itself below a chart is not a summary.
  const kpis = [tile(
    nextId(), `${profile.rowCount.toLocaleString()} rows`, 'kpi',
    { x: null, y: [{ column: null, agg: 'count' }] },
    { size: 'S', sort: { by: 'y', dir: 'desc' } },
  )];

  const rankedMeasures = measures
    .slice()
    .sort((a, b) => b.nonNullRatio - a.nonNullRatio || a.index - b.index);

  for (const measure of rankedMeasures.slice(0, MAX_MEASURE_KPIS)) {
    kpis.push(tile(
      nextId(), `Total ${measure.name}`, 'kpi',
      { x: null, y: [{ column: measure.name, agg: 'sum' }] },
      { size: 'S', isPercent: measure.isPercent },
    ));
  }

  // --- scored candidates ---------------------------------------------
  const candidates = [];

  if (primaryTemporal && topMeasure) {
    const bin = Number.isFinite(primaryTemporal.min) && Number.isFinite(primaryTemporal.max)
      ? chooseTruncation(primaryTemporal.min, primaryTemporal.max)
      : 'day';
    candidates.push({
      score: 1.2 * primaryTemporal.nonNullRatio,
      spec: tile(
        nextId(), `${topMeasure.name} over ${primaryTemporal.name}`, 'line',
        {
          x: { column: primaryTemporal.name, bin },
          y: [{ column: topMeasure.name, agg: 'sum' }],
        },
        // Chronological, always. A time series sorted by value is not a
        // time series.
        { sort: { by: 'x', dir: 'asc' }, limit: 200, size: 'L' },
      ),
    });
  }

  for (const dimension of chartableDimensions) {
    // With no measure to sum, counting rows is still a real chart --
    // "how many of each" is the question most sheets without numbers
    // are actually asking.
    const measured = topMeasure
      ? { title: `${topMeasure.name} by ${dimension.name}`, y: { column: topMeasure.name, agg: 'sum' } }
      : { title: `Rows by ${dimension.name}`, y: { column: null, agg: 'count' } };

    candidates.push({
      score: dimensionScore(dimension) * (1 + FOCUS_BONUS * relevanceOf(dimension.name, focus)),
      spec: tile(
        nextId(), measured.title, 'bar',
        { x: { column: dimension.name }, y: [measured.y] },
        { sort: { by: 'y', dir: 'desc' }, limit: 10 },
      ),
    });
  }

  for (const measure of measures) {
    candidates.push({
      // Below the paired charts: a distribution is useful but less
      // immediately readable than "X by Y".
      score: 0.6 * measureScore(measure),
      spec: tile(
        nextId(), `Distribution of ${measure.name}`, 'bar',
        {
          x: { column: measure.name, bin: 'auto' },
          y: [{ column: measure.name, agg: 'count' }],
        },
        { sort: { by: 'x', dir: 'asc' }, limit: 60 },
      ),
    });
  }

  if (dataset) {
    for (let a = 0; a < measures.length; a++) {
      for (let b = a + 1; b < measures.length; b++) {
        const ca = dataset.columns[dataset.byName.get(measures[a].name)];
        const cb = dataset.columns[dataset.byName.get(measures[b].name)];
        if (!ca || !cb) continue;
        const r = correlation(ca.values, cb.values, dataset.rowCount);
        if (Math.abs(r) < SCATTER_MIN_R) continue;
        candidates.push({
          score: Math.abs(r),
          spec: tile(
            nextId(), `${measures[b].name} against ${measures[a].name}`, 'scatter',
            {
              x: { column: measures[a].name },
              y: [
                { column: measures[a].name, agg: 'sum' },
                { column: measures[b].name, agg: 'sum' },
              ],
            },
            { sort: { by: 'x', dir: 'asc' }, limit: 500 },
          ),
        });
      }
    }
  }

  candidates.sort((x, y) => y.score - x.score);

  return [...kpis, ...candidates.map((c) => c.spec)].slice(0, MAX_TILES);
}
