// Which columns hold prose worth reading -- spec §6.1.
//
// This deliberately does NOT gate on the profiler's `text` verdict, and
// that is the whole subtlety of the module. The profiler calls a column
// categorical whenever it has 50 or fewer distinct values, so on a
// survey with 42 responses the written-answer column comes back
// `categorical` -- and gating on `text` would mean this feature never
// fired on precisely the surveys it was built for. Measured on the real
// export before the rule was written.
//
// What actually separates prose from a category is shape, not type:
//
//   * length     -- a column of 70-character sentences is somebody
//                   writing; a column of 15-character values is a code
//   * uniqueness -- everybody phrases a complaint differently; a
//                   category is chosen from a short list and repeats
//   * fill       -- a column mostly empty is an optional note, not the
//                   question being analysed
//
// Types that cannot be prose whatever their shape are excluded outright.

import { normalizeText } from './boilerplate.js';

export const MIN_MEAN_LENGTH = 40;
export const MIN_FILLED_RATIO = 0.6;
export const MIN_DISTINCT_RATIO = 0.8;

// A number, a date, a yes/no, an identifier or a multi-select is never
// free text, however long it happens to be.
const NEVER_PROSE = new Set([
  'numeric', 'date', 'datetime', 'boolean', 'identifier', 'multi', 'empty',
]);

export function detectTextColumns(profile, grid) {
  const rows = grid?.rows ?? [];
  const found = [];

  for (const column of profile?.columns ?? []) {
    if (NEVER_PROSE.has(column.type)) continue;

    const values = rows.map((row) => normalizeText(row?.[column.index]));
    const filled = values.filter((v) => v !== '');
    if (filled.length === 0) continue;
    if (filled.length / values.length < MIN_FILLED_RATIO) continue;

    const distinct = new Set(filled).size;
    if (distinct / filled.length < MIN_DISTINCT_RATIO) continue;

    const meanLength = filled.reduce((sum, v) => sum + v.length, 0) / filled.length;
    if (meanLength < MIN_MEAN_LENGTH) continue;

    found.push({ name: column.name, index: column.index, meanLength });
  }

  // Longest first: on a survey with two free-text questions, the one
  // people wrote most in is the one they were asked to describe.
  return found.sort((a, b) => b.meanLength - a.meanLength);
}
