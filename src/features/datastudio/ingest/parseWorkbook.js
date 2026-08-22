// SheetJS wrapper: an uploaded file's bytes -> raw row-major grids.
//
// Every option here is load-bearing; the defaults are wrong for us.

import { read, utils } from 'xlsx';

export function parseWorkbook(arrayBuffer) {
  const wb = read(arrayBuffer, {
    type: 'array',
    // Real Excel date cells arrive as `Date` objects rather than serial
    // numbers or locale-formatted strings. That removes D/M/Y ambiguity
    // entirely for them (spec §7.5) -- the only dates we have to guess
    // at are the ones typed as text.
    cellDates: true,
    cellNF: false,
    cellText: false,
  });

  return {
    sheets: wb.SheetNames.map((name) => ({
      name,
      rows: utils.sheet_to_json(wb.Sheets[name], {
        // Row-major arrays, not objects. Object mode keys rows by header
        // name and so silently drops every duplicate column -- and
        // duplicate columns are exactly what `toGrid` exists to
        // disambiguate.
        header: 1,
        raw: true,
        defval: null,
        blankrows: false,
      }),
    })),
  };
}
