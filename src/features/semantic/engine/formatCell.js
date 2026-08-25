// One stored cell -> the text a person should read.
//
// The columnar store keeps integers: a category is a dictionary
// position, a date is epoch milliseconds, a boolean is 0/1/2. Nothing
// in it is readable on its own, so every surface that shows a VALUE
// rather than an aggregate has to decode it the same way -- the CSV
// export, the row table, the record card. Having each do its own
// decoding is how they drift, and the drift is invisible: a wrong
// label still looks like a label.
//
// The order of the branches is load-bearing. A `multi` column carries a
// dictionary too, so testing `column.dictionary` first reads its FLAT
// option array as one code per row and prints whatever option happens
// to sit at that offset -- a confident, wrong answer. Multi comes
// first, exactly as it does in `aggregate.js` and `filterMask.js`.

import { formatMYT } from '../../../utils/malaysiaTime.js';

export const EMPTY_CELL = '';

/**
 * The decoded text of one cell, or `EMPTY_CELL` when it is missing.
 *
 * Missing is never invented: a blank number does not become 0 and a
 * blank category does not become the first category. The null
 * encodings this tests for are the contract set in `dataset.js`.
 *
 * `percentAsText` is what separates reading from exporting. On screen a
 * ratio column should read "12.5%"; in a CSV bound for Excel it must
 * stay the number 0.125, or every one of those cells arrives as text
 * and no formula will touch it.
 */
export function formatCell(column, index, { percentAsText = true } = {}) {
  if (!column) return EMPTY_CELL;

  // A multi column's row is a RANGE of the flat array, not one slot.
  // An empty range is the null encoding for this type.
  if (column.type === 'multi') {
    const offsets = column.offsets;
    if (!offsets) return EMPTY_CELL;
    const start = offsets[index];
    const end = offsets[index + 1];
    if (!(end > start)) return EMPTY_CELL;
    const labels = [];
    for (let i = start; i < end; i++) {
      const label = column.dictionary?.[column.values[i]];
      if (label !== undefined) labels.push(label);
    }
    return labels.join(', ');
  }

  const raw = column.values[index];

  if (column.dictionary) {
    return raw < 0 ? EMPTY_CELL : (column.dictionary[raw] ?? EMPTY_CELL);
  }

  if (column.type === 'boolean') {
    if (raw === 2) return EMPTY_CELL;
    return raw === 1 ? 'Yes' : 'No';
  }

  if (column.type === 'date' || column.type === 'datetime') {
    if (Number.isNaN(raw)) return EMPTY_CELL;
    // A date-only column has no time of day, so rendering 00:00 against
    // every row would invent precision the data does not have.
    return formatMYT(raw, column.dateOnly ? 'date' : 'datetime');
  }

  if (typeof raw === 'number') {
    if (Number.isNaN(raw)) return EMPTY_CELL;
    if (column.isPercent && percentAsText) return `${(raw * 100).toFixed(1)}%`;
    return String(raw);
  }

  return raw ?? EMPTY_CELL;
}
