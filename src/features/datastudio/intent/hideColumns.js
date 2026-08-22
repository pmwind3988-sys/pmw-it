// Parking a column, and letting it back.
//
// Hiding is a ROLE change and nothing else. The column keeps its type,
// its stats and its values; it simply stops being offered as an axis, a
// measure or a filter -- which is exactly what role `ignored` already
// means everywhere else in Data Studio, so nothing downstream needs a
// second concept for "the autopilot put this aside".
//
// The role is changed in place rather than by re-profiling, because
// re-profiling a hidden column would cost a pass over every row to
// arrive at the stats it already has. `overrideColumn` in the provider
// still re-profiles, and must: that one is the user changing a column's
// TYPE, where the stats genuinely differ.

import { ROLE_BY_TYPE } from '../profile/profileColumn.js';
import { retopProfile } from '../profile/profileDataset.js';

export function hideColumns(profile, hidden = []) {
  if (!profile || hidden.length === 0) return profile;
  const names = new Set(hidden.map((c) => c.name));

  return retopProfile({
    ...profile,
    columns: profile.columns.map((column) => (names.has(column.name)
      // `overridden` is what the profile panel reads to show a column as
      // decided rather than inferred, and this decision is as much an
      // override as one the user typed.
      ? { ...column, role: 'ignored', overridden: true }
      : column)),
  });
}

/**
 * Give the hidden columns their inferred role back.
 *
 * `ROLE_BY_TYPE` rather than a remembered role: the type is unchanged,
 * so the role the profiler would have given it is recoverable from the
 * type alone, and storing the old role would be a second copy of the
 * same fact waiting to disagree with the first.
 */
export function unhideColumns(profile, hidden = []) {
  if (!profile || hidden.length === 0) return profile;
  const names = new Set(hidden.map((c) => c.name));

  return retopProfile({
    ...profile,
    columns: profile.columns.map((column) => (names.has(column.name)
      ? { ...column, role: ROLE_BY_TYPE[column.type] ?? 'ignored', overridden: false }
      : column)),
  });
}
