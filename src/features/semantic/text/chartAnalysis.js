// What a finished reading does to the sheet: five more columns, and the
// charts that draw them.
//
// A plain state -> state function rather than a method on the provider,
// because it runs from two places -- by itself the moment a reading
// lands, and again from a button once the user has corrected something
// -- and because everything it decides is worth testing without React
// in the way.
//
// Running it twice must REPLACE the analysis columns and their charts,
// never stack a second copy on top. That is what dropping
// `DERIVED_HEADERS` before appending, and the stable tile ids in
// `analysisTiles.js`, are for.

import { proposeCleanPlan } from '../clean/proposeCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';
import { detectTextColumns } from './detectTextColumns.js';
import { deriveColumns, DERIVED_OVERRIDES, DERIVED_HEADERS } from './deriveColumns.js';
import { withAnalysisTiles } from './analysisTiles.js';

/**
 * Appends the analysis to the grid and charts it.
 *
 * The grid is re-profiled afterwards rather than patched, so the new
 * columns go through the same type inference as every other column and
 * nothing downstream needs a special case. `DERIVED_OVERRIDES` supplies
 * the one thing inference cannot know: the categories column is
 * multi-valued by construction, even though most respondents raise a
 * single category and most of its cells therefore carry no separator.
 *
 * Returns the state unchanged when there is nothing to chart, so the
 * caller never has to check first.
 */
export function withAnalysisCharted(state) {
  if (!state?.analysis || !state?.grid) return state;

  const { headers, columns } = deriveColumns(state.analysis, state.grid.rows.length);

  const keep = state.grid.headers
    .map((name, i) => ({ name, i }))
    .filter(({ name }) => !DERIVED_HEADERS.includes(name));

  const grid = {
    headers: [...keep.map((k) => k.name), ...headers],
    rows: state.grid.rows.map((row, r) => [
      ...keep.map(({ i }) => row?.[i] ?? null),
      ...columns.map((column) => column[r]),
    ]),
  };
  const profile = profileDataset(grid, DERIVED_OVERRIDES);

  return {
    ...state,
    grid,
    profile,
    plan: proposeCleanPlan(profile, grid),
    textColumns: detectTextColumns(profile, grid),
    tiles: withAnalysisTiles(state.tiles, headers),
    autoCharted: true,
  };
}
