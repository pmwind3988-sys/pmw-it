/**
 * Where a machine physically is: F1, F3, PML and whatever else IT types.
 *
 * Stored as an upper-case CODE in a text column, never a choice column: a new
 * site must not need a SharePoint column change before a machine can be put
 * there. Which locations exist is read the way `assets/categories.js` reads
 * categories -- the built-in ones, then every one a row is actually using --
 * so there is no second list to disagree with the rows.
 */
export const BUILT_IN_LOCATIONS = ['F1', 'F3', 'PML'];

export function cleanLocation(typed) {
  const code = String(typed ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
  return code || null;
}

export function locationsIn(devices = []) {
  const list = [...BUILT_IN_LOCATIONS];
  const seen = new Set(list);
  for (const device of devices) {
    const code = cleanLocation(device?.location);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    list.push(code);
  }
  return list;
}
