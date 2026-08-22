// Finding the real header row in a sheet -- spec §2.1, §12.
//
// Exported spreadsheets rarely start with their header. A report title
// sits on row 1, a blank spacer on row 2, and the column names appear on
// row 3. Taking row 1 on faith gives every column a name like
// "IT Request Report 2024" and shifts every value up one row, which
// looks like a data problem rather than a parsing one.
//
// So we score candidate rows instead of trusting position, on three
// signals that together separate a header from its data:
//
//   1. Header cells are words, not values. "Amount" is a header;
//      "1,234" is not, however string-shaped it looks in the file.
//   2. A header row is fuller than the decoration above it. A title
//      banner fills one cell of twelve; the header fills all twelve.
//   3. The row BELOW a header looks different from it. Strings over
//      numbers is the single strongest signal a sheet gives us.

import { isNullish, parseNumberLike, parseBooleanLike } from '../profile/inferType.js';

// Scanning the whole sheet would be pointless -- a header buried below
// row 20 is not a header, it is a second table, which is out of scope.
export const HEADER_SCAN_ROWS = 20;

// Below this, no row is a convincing header and we say so rather than
// promoting the least-bad one and mislabelling every column.
export const HEADER_SCORE_FLOOR = 0.5;

// Relative weight of each signal. Type difference from the row below is
// weighted highest because it is the hardest to produce by accident.
const WEIGHT_WORDS = 0.4;
const WEIGHT_FILL = 0.25;
const WEIGHT_DIFFERS = 0.35;

// The coarse "kind" of a cell, which is all header detection needs. A
// numeric string is a number here: what matters is whether the cell
// reads as a value or as a label, not how it happens to be stored.
function cellShape(value) {
  if (value instanceof Date) return 'date';
  if (isNullish(value)) return 'empty';
  if (typeof value === 'number') return 'number';
  if (typeof value === 'boolean') return 'boolean';
  if (parseNumberLike(value).ok) return 'number';
  if (parseBooleanLike(value).ok) return 'boolean';
  return 'text';
}

function rowWidth(rows) {
  return rows.reduce((widest, row) => Math.max(widest, row?.length ?? 0), 0);
}

// How header-like one row is, in 0..1.
function scoreRow(rows, index, width) {
  const row = rows[index] ?? [];
  const shapes = Array.from({ length: width }, (_, c) => cellShape(row[c]));
  const filled = shapes.filter((s) => s !== 'empty');

  // A blank row is a spacer, never a header. Returning 0 rather than
  // letting it through on fill ratio alone keeps blank rows from
  // out-scoring a real header in a sparse sheet.
  if (filled.length === 0) return 0;

  const words = filled.filter((s) => s === 'text').length / filled.length;
  const fill = filled.length / width;

  // Compare against the first non-blank row below, not merely the next
  // one -- a blank spacer between the header and the data would
  // otherwise read as "totally different" and flatter every row above it.
  let differs = 0;
  for (let below = index + 1; below < rows.length; below++) {
    const belowShapes = Array.from({ length: width }, (_, c) => cellShape(rows[below]?.[c]));
    if (belowShapes.every((s) => s === 'empty')) continue;
    let changed = 0;
    for (let c = 0; c < width; c++) {
      if (shapes[c] !== belowShapes[c]) changed++;
    }
    differs = changed / width;
    break;
  }

  return WEIGHT_WORDS * words + WEIGHT_FILL * fill + WEIGHT_DIFFERS * differs;
}

export function detectHeader(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return { headerIndex: -1, confidence: 0 };
  }

  const width = rowWidth(rows);
  if (width === 0) return { headerIndex: -1, confidence: 0 };

  const limit = Math.min(rows.length, HEADER_SCAN_ROWS);

  let bestIndex = -1;
  let bestScore = 0;
  for (let i = 0; i < limit; i++) {
    const score = scoreRow(rows, i, width);
    // Strictly greater, so a tie keeps the earliest row: if two rows are
    // equally header-like, the one above them is the header and the one
    // below is a repeat of it inside the data.
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  if (bestScore < HEADER_SCORE_FLOOR) return { headerIndex: -1, confidence: bestScore };
  return { headerIndex: bestIndex, confidence: bestScore };
}

// Every column needs a name that is unique and non-empty, because names
// are how tiles, filters and saved dashboards refer to columns. A blank
// name would render as a nameless axis; a repeated one would make two
// different columns indistinguishable in a saved tile spec.
function nameColumns(headerRow, width) {
  const used = new Map();
  const names = [];

  for (let c = 0; c < width; c++) {
    const raw = headerRow?.[c];
    const base = isNullish(raw)
      ? `Column ${c + 1}`
      : String(raw instanceof Date ? raw.toISOString() : raw).trim();

    const seen = used.get(base) ?? 0;
    used.set(base, seen + 1);
    // First occurrence keeps the plain name; later ones are numbered
    // from 2, so a user reading "Name (3)" knows it is the third.
    names.push(seen === 0 ? base : `${base} (${seen + 1})`);
  }

  return names;
}

export function toGrid(rows, headerIndex) {
  const all = Array.isArray(rows) ? rows : [];
  const width = rowWidth(all);

  // No header found: keep every row as data and name the columns by
  // position, so the user still sees their data and can pick the header
  // manually rather than being handed an error.
  if (headerIndex < 0) {
    return { headers: nameColumns(null, width), rows: all };
  }

  return {
    headers: nameColumns(all[headerIndex], width),
    rows: all.slice(headerIndex + 1),
  };
}
