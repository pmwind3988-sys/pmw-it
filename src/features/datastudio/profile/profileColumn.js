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
import { toEpochMs } from '../time/malaysiaTime.js';

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

// A multi column's frequent values are its OPTIONS, not its cells. A
// cell reading "A;B;C" is three answers; counting it whole makes every
// respondent their own category and the filter picker useless.
function rankTopOptions(values, separator) {
  const counts = new Map();
  for (const v of values) {
    if (isNullish(v)) continue;
    for (const part of String(v).split(separator)) {
      const label = part.trim();
      if (label === '') continue;
      counts.set(label, (counts.get(label) ?? 0) + 1);
    }
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

// Which role each type plays on the canvas -- spec section 7.1. Kept as
// data rather than a chain of ifs because the profile panel's type
// override needs to answer the same question for a type the inference
// never picked.
export const ROLE_BY_TYPE = {
  numeric: 'measure',
  categorical: 'dimension',
  multi: 'dimension',
  boolean: 'dimension',
  date: 'temporal',
  datetime: 'temporal',
  text: 'ignored',
  identifier: 'ignored',
  empty: 'ignored',
};

// `override` is the user disagreeing with the inference from the profile
// panel, and the user wins -- that is the whole point of showing them the
// verdict. The measured confidence is kept as reported rather than forced
// to 1: it still describes how well the data fitted the ORIGINAL guess,
// which is the context for why they had to intervene.
function applyOverride(verdict, override) {
  if (!override?.type && !override?.role) return verdict;
  const type = override.type ?? verdict.type;
  const role = override.role ?? ROLE_BY_TYPE[type] ?? 'ignored';
  return { ...verdict, type, role, overridden: true };
}

// min/max for a temporal column, as epoch milliseconds.
//
// Without this the suggestion engine cannot see how long a date column
// spans, so `chooseTruncation` has nothing to choose from and every time
// series falls back to a day grain -- which on five years of data draws
// about eighteen hundred categories.
function temporalStats(values, dateOrder) {
  const order = dateOrder && dateOrder !== 'conflict' ? dateOrder : 'dmy';
  let min = null;
  let max = null;

  for (const v of values) {
    if (isNullish(v)) continue;
    const ms = v instanceof Date ? v.getTime() : toEpochMs(v, { order, dateOnly: true });
    if (!Number.isFinite(ms)) continue;
    if (min === null || ms < min) min = ms;
    if (max === null || ms > max) max = ms;
  }

  return { min, max, mean: null };
}

export function profileColumn(values, columnName, index, override = null) {
  const verdict = applyOverride(inferType(values, columnName), override);
  const total = values.length;

  // An empty grid has no rows to be non-null, and 0/0 is NaN -- which
  // would then propagate into topMeasure ranking and silently lose every
  // comparison. Treat "no rows at all" as a ratio of 0.
  const nonNullRatio = total > 0 ? (total - verdict.nullCount) / total : 0;

  const isNumericLike = verdict.type === 'numeric';
  const isTemporalLike = verdict.type === 'date' || verdict.type === 'datetime';
  const isDimensionLike = verdict.role === 'dimension';

  return {
    ...verdict,
    name: columnName,
    index,
    nonNullRatio,
    topValues: verdict.type === 'multi'
      ? rankTopOptions(values, verdict.separator ?? ';')
      : (isDimensionLike ? rankTopValues(values) : []),
    ...(isNumericLike ? numericStats(values) : null),
    ...(isTemporalLike ? temporalStats(values, verdict.dateOrder) : null),
    ...(isNumericLike || isTemporalLike ? null : { min: null, max: null, mean: null }),
  };
}
