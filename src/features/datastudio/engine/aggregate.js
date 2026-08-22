// Group-by, measures, binning, date truncation and top-N -- spec §10.2.
//
// One pass over the masked rows accumulates every (category, series)
// pair into a Map, and only then is the dense matrix materialised. The
// alternative -- a pass per series -- multiplies the cost by the series
// count for no benefit.
//
// Two rules here are the ones that produce wrong numbers when broken,
// so both are pinned by tests:
//
//   * `avg` divides by the count of NON-NULL values, not by the group
//     size. Dividing by the group size quietly treats every missing
//     value as a zero and drags the average down.
//   * `Other` is always last. Sorted into the middle by its own value it
//     stops reading as "everything else" and reads as a category.

const DAY_MS = 86400000;

// Spec §10.5's thresholds for picking a time grain. Below 90 days a day
// grain is readable; below 3 years a month grain is; beyond that,
// quarters.
const DAY_GRAIN_LIMIT_MS = 90 * DAY_MS;
const MONTH_GRAIN_LIMIT_MS = 3 * 365 * DAY_MS;

export const OTHER_LABEL = 'Other';

// A guard, not a preference: Freedman-Diaconis can propose an enormous
// bin count on a spiky distribution, and a chart with 4000 bars is a
// solid rectangle.
const MAX_BINS = 60;

export function truncateDate(epochMs, unit) {
  const d = new Date(epochMs);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth();

  switch (unit) {
    case 'year':
      return Date.UTC(y, 0, 1);
    case 'quarter':
      return Date.UTC(y, Math.floor(m / 3) * 3, 1);
    case 'month':
      return Date.UTC(y, m, 1);
    case 'day':
    default:
      return Date.UTC(y, m, d.getUTCDate());
  }
}

export function chooseTruncation(minMs, maxMs) {
  const span = Math.abs((maxMs ?? 0) - (minMs ?? 0));
  if (span < DAY_GRAIN_LIMIT_MS) return 'day';
  if (span < MONTH_GRAIN_LIMIT_MS) return 'month';
  return 'quarter';
}

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Labels are built from the UTC parts directly rather than through Intl.
// The epochs are already truncated to UTC boundaries, so re-interpreting
// them in a named zone would only reintroduce the off-by-one-day risk
// that truncating removed.
function temporalLabel(epochMs, unit) {
  const d = new Date(epochMs);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth();
  switch (unit) {
    case 'year':
      return String(y);
    case 'quarter':
      return `Q${Math.floor(m / 3) + 1} ${y}`;
    case 'month':
      return `${MONTH_NAMES[m]} ${y}`;
    case 'day':
    default:
      return `${String(d.getUTCDate()).padStart(2, '0')}/${String(m + 1).padStart(2, '0')}/${y}`;
  }
}

function niceNumber(n) {
  if (!Number.isFinite(n)) return '';
  const abs = Math.abs(n);
  if (abs >= 1000 || abs === 0 || abs >= 1) return String(Math.round(n * 100) / 100);
  return String(Number(n.toPrecision(3)));
}

/**
 * Freedman-Diaconis binning: width = 2 * IQR / n^(1/3).
 *
 * FD is used rather than a fixed bin count because it adapts to spread:
 * a tight distribution gets fine bins and a sprawling one gets coarse
 * ones, which is what makes a histogram of unknown data readable
 * without the user tuning anything.
 */
export function binNumeric(values, mask) {
  const kept = [];
  for (let i = 0; i < values.length; i++) {
    if (mask && !mask[i]) continue;
    const v = values[i];
    if (Number.isNaN(v)) continue;
    kept.push(v);
  }

  if (kept.length === 0) return { edges: [], labels: [] };

  kept.sort((a, b) => a - b);
  const min = kept[0];
  const max = kept[kept.length - 1];

  // Every value identical: FD's width is 0 and the bin count is
  // infinite. One bin is the honest answer.
  if (min === max) {
    return { edges: [min, min], labels: [niceNumber(min)] };
  }

  const q = (p) => kept[Math.min(kept.length - 1, Math.floor(p * (kept.length - 1)))];
  const iqr = q(0.75) - q(0.25);

  // A zero IQR with a non-zero range means most values are identical
  // with a few outliers. FD gives no usable width there, so fall back to
  // Sturges, which only needs the count.
  const width = iqr > 0
    ? (2 * iqr) / Math.cbrt(kept.length)
    : (max - min) / Math.ceil(Math.log2(kept.length) + 1);

  const binCount = Math.max(
    1,
    Math.min(MAX_BINS, Math.ceil((max - min) / (width > 0 ? width : (max - min)))),
  );
  const step = (max - min) / binCount;

  const edges = Array.from({ length: binCount + 1 }, (_, i) => min + i * step);
  // Nudge the last edge outward so the maximum value falls inside the
  // final bin rather than off its right-hand end.
  edges[binCount] = max;

  const labels = Array.from({ length: binCount }, (_, i) => (
    `${niceNumber(edges[i])}–${niceNumber(edges[i + 1])}`));

  return { edges, labels };
}

// --- grouping ---------------------------------------------------------

function columnOf(dataset, name) {
  const index = dataset?.byName?.get(name);
  return index === undefined ? undefined : dataset.columns[index];
}

// The label a row's x value groups under, plus a sort key. Returns null
// when the row has no x value at all -- a row that cannot be placed on
// the axis is not a category called "empty", it is a row this chart
// cannot show.
function makeXResolver(dataset, xSpec, mask) {
  const column = columnOf(dataset, xSpec?.column);
  if (!column) return null;

  const isTemporal = column.type === 'date' || column.type === 'datetime';
  const bin = xSpec?.bin;

  if (isTemporal && bin) {
    let unit = bin;
    if (bin === 'auto') {
      let min = Infinity;
      let max = -Infinity;
      for (let i = 0; i < column.values.length; i++) {
        if (mask && !mask[i]) continue;
        const v = column.values[i];
        if (Number.isNaN(v)) continue;
        if (v < min) min = v;
        if (v > max) max = v;
      }
      unit = Number.isFinite(min) ? chooseTruncation(min, max) : 'day';
    }
    return (row) => {
      const v = column.values[row];
      if (Number.isNaN(v)) return null;
      const key = truncateDate(v, unit);
      return { key, label: temporalLabel(key, unit) };
    };
  }

  if (column.type === 'numeric' && bin) {
    const { edges, labels } = binNumeric(column.values, mask);
    if (labels.length === 0) return () => null;
    return (row) => {
      const v = column.values[row];
      if (Number.isNaN(v)) return null;
      // The last bin is closed at both ends so the maximum value has
      // somewhere to go.
      let index = labels.length - 1;
      for (let b = 0; b < labels.length; b++) {
        if (v < edges[b + 1]) {
          index = b;
          break;
        }
      }
      return { key: index, label: labels[index] };
    };
  }

  if (column.dictionary) {
    return (row) => {
      const code = column.values[row];
      if (code < 0) return null;
      return { key: code, label: column.dictionary[code] };
    };
  }

  if (column.type === 'boolean') {
    return (row) => {
      const v = column.values[row];
      if (v === 2) return null;
      return { key: v, label: v === 1 ? 'Yes' : 'No' };
    };
  }

  if (isTemporal || column.type === 'numeric') {
    return (row) => {
      const v = column.values[row];
      if (Number.isNaN(v)) return null;
      return {
        key: v,
        label: isTemporal ? temporalLabel(truncateDate(v, 'day'), 'day') : niceNumber(v),
      };
    };
  }

  return (row) => {
    const v = column.values[row];
    if (v === null || v === undefined) return null;
    return { key: v, label: String(v) };
  };
}

function makeSeriesResolver(dataset, seriesSpec) {
  const column = columnOf(dataset, seriesSpec?.column);
  if (!column) return null;

  if (column.dictionary) {
    return (row) => {
      const code = column.values[row];
      return code < 0 ? null : column.dictionary[code];
    };
  }
  if (column.type === 'boolean') {
    return (row) => {
      const v = column.values[row];
      return v === 2 ? null : (v === 1 ? 'Yes' : 'No');
    };
  }
  return (row) => {
    const v = column.values[row];
    return v === null || v === undefined ? null : String(v);
  };
}

function newBucket() {
  // `count` is rows; `values` is only the non-null measure values, which
  // is what every aggregation except `count` reads. Keeping them apart
  // is what makes avg divide by the right denominator.
  return { count: 0, sum: 0, values: [] };
}

function reduceBucket(bucket, agg) {
  const { values } = bucket;

  switch (agg) {
    case 'count':
      return bucket.count;
    case 'countDistinct':
      return new Set(values).size;
    case 'sum':
      return bucket.sum;
    case 'avg':
      // Non-null denominator. Dividing by `bucket.count` would treat
      // every missing value as a zero.
      return values.length > 0 ? bucket.sum / values.length : 0;
    case 'min':
      return values.length > 0 ? Math.min(...values) : 0;
    case 'max':
      return values.length > 0 ? Math.max(...values) : 0;
    case 'median': {
      if (values.length === 0) return 0;
      const sorted = values.slice().sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 === 0
        ? (sorted[mid - 1] + sorted[mid]) / 2
        : sorted[mid];
    }
    default:
      return bucket.sum;
  }
}

export function aggregate(dataset, mask, spec) {
  const measure = spec?.encoding?.y?.[0] ?? {};
  const agg = measure.agg ?? 'count';
  const measureColumn = columnOf(dataset, measure.column);
  const seriesName = spec?.encoding?.series?.column
    ? null
    : (measure.column ?? 'Count');

  const resolveX = makeXResolver(dataset, spec?.encoding?.x, mask);
  const resolveSeries = makeSeriesResolver(dataset, spec?.encoding?.series);

  const empty = { categories: [], series: [{ name: seriesName ?? 'Count', data: [] }] };
  if (!resolveX) return empty;

  // key -> { label, sortKey, bySeries: Map<seriesName, bucket>, total }
  const groups = new Map();
  const seriesNames = [];
  const seenSeries = new Set();

  const rowCount = dataset.rowCount;
  for (let row = 0; row < rowCount; row++) {
    if (mask && !mask[row]) continue;

    const x = resolveX(row);
    if (x === null) continue;

    let group = groups.get(x.key);
    if (!group) {
      group = { label: x.label, sortKey: x.key, bySeries: new Map(), total: 0 };
      groups.set(x.key, group);
    }

    const name = resolveSeries ? resolveSeries(row) : seriesName;
    if (name === null) continue;
    if (!seenSeries.has(name)) {
      seenSeries.add(name);
      seriesNames.push(name);
    }

    let bucket = group.bySeries.get(name);
    if (!bucket) {
      bucket = newBucket();
      group.bySeries.set(name, bucket);
    }

    bucket.count++;
    if (measureColumn) {
      const v = measureColumn.values[row];
      if (typeof v === 'number' && !Number.isNaN(v)) {
        bucket.sum += v;
        bucket.values.push(v);
      }
    }
  }

  if (groups.size === 0) return empty;

  // Reduce every bucket before sorting -- the sort compares final
  // values, and for `avg` or `median` those are nothing like the sums.
  const rows = [...groups.values()].map((group) => {
    const byName = new Map();
    let total = 0;
    for (const [name, bucket] of group.bySeries) {
      const value = reduceBucket(bucket, agg);
      byName.set(name, value);
      total += value;
    }
    return { label: group.label, sortKey: group.sortKey, byName, total };
  });

  const dir = spec?.sort?.dir === 'asc' ? 1 : -1;
  if (spec?.sort?.by === 'x') {
    rows.sort((a, b) => {
      if (typeof a.sortKey === 'number' && typeof b.sortKey === 'number') {
        return (a.sortKey - b.sortKey) * dir;
      }
      return String(a.label).localeCompare(String(b.label)) * dir;
    });
  } else {
    rows.sort((a, b) => (a.total - b.total) * dir);
  }

  // Top-N: keep `limit` rows and fold the rest into a single trailing
  // 'Other'. It is appended after the sort, never sorted, so it always
  // reads as "everything else" rather than as another category.
  const limit = spec?.limit ?? rows.length;
  let kept = rows;
  let other = null;

  if (Number.isFinite(limit) && rows.length > limit) {
    kept = rows.slice(0, limit);
    const rest = rows.slice(limit);
    const byName = new Map();
    for (const name of seriesNames) {
      let sum = 0;
      for (const r of rest) sum += r.byName.get(name) ?? 0;
      byName.set(name, sum);
    }
    other = { label: OTHER_LABEL, byName };
  }

  const finalRows = other ? [...kept, other] : kept;
  const names = seriesNames.length > 0 ? seriesNames : [seriesName ?? 'Count'];

  return {
    categories: finalRows.map((r) => r.label),
    series: names.map((name) => ({
      name,
      // Padded dense: a series with no rows in a category gets 0 there,
      // so every series array lines up with `categories` index for
      // index. A ragged array would silently shift a bar onto the wrong
      // label.
      data: finalRows.map((r) => r.byName.get(name) ?? 0),
    })),
  };
}
