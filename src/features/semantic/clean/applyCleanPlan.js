// Runs an approved clean plan and hands back the typed dataset.
//
// The input grid is never mutated. The whole non-destructive model in
// spec §8.1 rests on that: the user unticks a step, the plan is re-run
// from the raw grid, and the previous result leaves no residue. If any
// op wrote through to the source, "untick" would only ever mean "apply
// the remaining steps to already-damaged data".

import {
  trimWhitespace, normalizeNulls, parseNumber, unifyCase, mergeCategories,
  parseDate, castType, dropEmptyColumns, dropEmptyRows, dedupeRows,
} from './cleanOps.js';
import { buildDataset } from '../engine/dataset.js';

const COLUMN_OPS = {
  trimWhitespace,
  normalizeNulls,
  parseNumber,
  unifyCase,
  mergeCategories,
  parseDate,
  castType,
};

const GRID_OPS = {
  dropEmptyColumns,
  dropEmptyRows,
  dedupeRows,
};

function transpose(headers, rows) {
  const columns = headers.map(() => []);
  for (const row of rows) {
    for (let c = 0; c < headers.length; c++) columns[c].push(row?.[c]);
  }
  return columns;
}

export function applyCleanPlan(grid, plan, profile) {
  const enabled = (plan ?? []).filter((s) => s.enabled);

  // Whole-grid ops run first, against the row-major form.
  //
  // This is deliberately NOT the order the plan lists them in. The
  // checklist shows them last because that reads naturally ("...and
  // finally, drop the empty rows"), but every count the user approved
  // was measured against the RAW grid -- "remove 3 duplicate rows" was
  // 3 duplicates of the untrimmed values. Applying them to the raw grid
  // is what makes the number they agreed to the number they get.
  let headers = grid.headers.slice();
  let rows = grid.rows.map((row) => (row ?? []).slice());

  for (const stepSpec of enabled) {
    const op = GRID_OPS[stepSpec.op];
    if (!op) continue;
    ({ headers, rows } = op({ headers, rows }));
  }

  // Then per-column ops, in plan order, so each sees the previous one's
  // output: trimming before merging before casting.
  const columns = transpose(headers, rows);
  const indexByName = new Map(headers.map((name, i) => [name, i]));

  for (const stepSpec of enabled) {
    const op = COLUMN_OPS[stepSpec.op];
    if (!op) continue;
    const index = indexByName.get(stepSpec.column);
    // A step naming a column that a grid op just dropped is not an
    // error -- it is a step with nothing left to do.
    if (index === undefined) continue;
    columns[index] = op(columns[index], stepSpec.params ?? {});
  }

  // Only the columns that survived, in their surviving order.
  const survivingProfile = profile
    ? { ...profile, columns: profile.columns.filter((c) => indexByName.has(c.name)) }
    : profile;

  return buildDataset({ headers, columns, profile: survivingProfile });
}
