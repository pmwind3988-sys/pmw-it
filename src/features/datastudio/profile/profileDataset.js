// Profiles every column of a parsed grid -- spec §7.1, §10.5.
//
// The grid arrives row-major, the way a spreadsheet is read, but every
// profiling question ("is this column numeric?") is asked column-wise.
// Transposing once up front costs one pass; asking per column would walk
// all the rows again for every column, which at 100k rows x 40 columns
// is the difference between a moment and a hang.

import { profileColumn } from './profileColumn.js';

// Row-major grid -> one array of values per column.
function transpose(headers, rows) {
  const columns = headers.map(() => []);
  for (const row of rows) {
    for (let c = 0; c < headers.length; c++) {
      columns[c].push(row?.[c]);
    }
  }
  return columns;
}

// The column of `role` with the highest non-null ratio, ties broken by
// original column order (spec §10.5). Defined here once and imported
// wherever it is needed -- re-deriving it elsewhere is how two parts of
// the UI end up disagreeing about which measure is "the" measure.
function pickByRole(columns, role) {
  let best = null;
  for (const column of columns) {
    if (column.role !== role) continue;
    // Strictly greater, so an equal ratio leaves the earlier column in
    // place and the tie-break falls out of the iteration order.
    if (best === null || column.nonNullRatio > best.nonNullRatio) best = column;
  }
  return best ? best.name : null;
}

export function profileDataset(grid) {
  const headers = grid?.headers ?? [];
  const rows = grid?.rows ?? [];

  const byColumn = transpose(headers, rows);
  const columns = headers.map((name, index) =>
    profileColumn(byColumn[index], name, index));

  return {
    columns,
    rowCount: rows.length,
    topMeasure: pickByRole(columns, 'measure'),
    primaryTemporal: pickByRole(columns, 'temporal'),
  };
}
