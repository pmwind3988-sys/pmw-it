// One column's full profile -- spec §7.1.
//
// `inferType` decides WHAT a column is; this decides what we can say
// ABOUT it. The verdict is passed through untouched and the descriptive
// statistics the canvas needs are layered on top: how full the column
// is, its most common values (so a dimension can be previewed and
// filtered without scanning the raw grid again), and min/max/mean for
// measures.
//
// Statistics are only computed where they mean something. A mean over a
// categorical column is not a smaller number than a mean over a measure
// -- it is nonsense -- so non-numeric columns carry `null` rather than a
// zero that a chart would happily plot.

import { inferType, isNullish, parseNumberLike } from './inferType.js';

// How many frequent values a dimension carries for previews and filter
// pickers. Spec §7.1; the profile panel shows fewer, but suggestion
// scoring wants the tail.
export const TOP_VALUE_LIMIT = 10;

function normalizeLabel(value) {
  if (value instanceof Date) return value.toISOString();
  return String(value ?? '').trim();
}

// Frequency of each non-null value, highest first. Ties keep the order
// the values were first seen in, so the result is stable across runs
// rather than depending on Map iteration incidentals.
function rankTopValues(values) {
  const counts = new Map();
  for (const v of values) {
    if (isNullish(v)) continue;
    const key = normalizeLabel(v);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return [...counts.entries()]
    .map(([value, count], firstSeen) => ({ value, count, firstSeen }))
    .sort((a, b) => b.count - a.count || a.firstSeen - b.firstSeen)
    .slice(0, TOP_VALUE_LIMIT)
    .map(({ value, count }) => ({ value, count }));
}

// min/max/mean over the values that actually parse as numbers. Values
// that do not parse are the column's casualties and are excluded rather
// than counted as zero, which would drag the mean toward nothing.
function numericStats(values) {
  let min = null;
  let max = null;
  let sum = 0;
  let count = 0;

  for (const v of values) {
    if (isNullish(v)) continue;
    const parsed = parseNumberLike(v);
    if (!parsed.ok) continue;
    const n = parsed.value;
    if (min === null || n < min) min = n;
    if (max === null || n > max) max = n;
    sum += n;
    count++;
  }

  return { min, max, mean: count > 0 ? sum / count : null };
}

export function profileColumn(values, columnName, index) {
  const verdict = inferType(values, columnName);
  const total = values.length;

  // An empty grid has no rows to be non-null, and 0/0 is NaN -- which
  // would then propagate into topMeasure ranking and silently lose every
  // comparison. Treat "no rows at all" as a ratio of 0.
  const nonNullRatio = total > 0 ? (total - verdict.nullCount) / total : 0;

  const isNumericLike = verdict.type === 'numeric';
  const isDimensionLike = verdict.role === 'dimension';

  return {
    ...verdict,
    name: columnName,
    index,
    nonNullRatio,
    topValues: isDimensionLike ? rankTopValues(values) : [],
    ...(isNumericLike
      ? numericStats(values)
      : { min: null, max: null, mean: null }),
  };
}
