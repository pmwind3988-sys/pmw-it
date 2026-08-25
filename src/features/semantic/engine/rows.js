// The rows behind the charts -- reading the store back a ROW at a time.
//
// Everything else in `engine/` asks column-shaped questions, because
// that is what charts need. This file is the one place that asks the
// opposite question: "which actual records am I looking at, and what
// does one of them say in full?" A dashboard that cannot answer that
// leaves the user with totals they have no way to check.
//
// Two rules keep it honest at 100k rows:
//
// 1. Nothing materialises the whole dataset. `pageRowIndexes` walks the
//    mask and stops as soon as it has filled one page, so the cost of
//    the panel is the cost of a page, not of the sheet.
// 2. Row order is source order. There is no sort, because a sort that
//    is not stable across a filter change moves the record somebody was
//    reading out from under them.

import { formatCell } from './formatCell.js';

/** How many of the dataset's rows the mask keeps. */
export function countRows(dataset, mask) {
  if (!dataset) return 0;
  if (!mask) return dataset.rowCount;
  let n = 0;
  for (let i = 0; i < mask.length; i++) if (mask[i]) n += 1;
  return n;
}

/**
 * The dataset row indexes on one page of the filtered set.
 *
 * Returns real indexes into the store, not positions within the page,
 * because everything downstream -- the detail card, the next/previous
 * buttons -- needs to address the row itself.
 */
export function pageRowIndexes(dataset, mask, offset = 0, limit = 50) {
  const out = [];
  if (!dataset || limit <= 0) return out;

  let seen = 0;
  for (let i = 0; i < dataset.rowCount; i++) {
    if (mask && !mask[i]) continue;
    if (seen >= offset) {
      out.push(i);
      if (out.length >= limit) break;
    }
    seen += 1;
  }
  return out;
}

/**
 * Steps to the next or previous row WITHIN the filtered set.
 *
 * Walking the raw index by one would wander into rows the current
 * filters exclude, so a user paging through "the HR records" would
 * silently be shown a Finance one. Returns null at either end.
 */
export function stepRow(dataset, mask, index, direction) {
  if (!dataset) return null;
  const step = direction < 0 ? -1 : 1;
  for (let i = index + step; i >= 0 && i < dataset.rowCount; i += step) {
    if (!mask || mask[i]) return i;
  }
  return null;
}

/**
 * One record in full: every column, decoded, in sheet order.
 *
 * Ignored columns are included and flagged rather than hidden. "All
 * the details" has to mean all of them -- the parked columns are
 * usually the timestamp and the respondent, which is exactly what
 * somebody drilling into a single answer is looking for.
 */
export function readRecord(dataset, index) {
  if (!dataset || index === null || index === undefined) return [];
  if (index < 0 || index >= dataset.rowCount) return [];

  return dataset.columns.map((column) => {
    const text = formatCell(column, index);
    return {
      name: column.name,
      type: column.type,
      role: column.role,
      parked: column.role === 'ignored',
      empty: text === '',
      text,
    };
  });
}

/** One page of the table, already decoded to text. */
export function readRows(dataset, mask, columns, offset = 0, limit = 50) {
  return pageRowIndexes(dataset, mask, offset, limit).map((index) => ({
    index,
    cells: columns.map((column) => formatCell(column, index)),
  }));
}
