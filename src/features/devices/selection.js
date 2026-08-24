/**
 * Which rows of the register are ticked, and what the header's tick box shows.
 *
 * Pure and set-based on purpose: the register is the one screen that can remove
 * several machines at once, and every rule about what is ticked -- above all
 * that a tick cannot survive off screen -- is a rule that can be wrong without
 * looking wrong.
 */

/**
 * A row with no SharePoint id cannot be removed (`deleteDevice` refuses it), so
 * it cannot be ticked either. Offering a tick box that leads to an error is
 * worse than offering none.
 */
export const isSelectable = (device) =>
  device !== null && device !== undefined && device.id !== null && device.id !== undefined;

/** The ids of everything the table is currently showing, in its own order. */
export const selectableIds = (rows) => rows.filter(isSelectable).map((row) => row.id);

export function toggleId(selected, id) {
  const next = new Set(selected);
  if (next.has(id)) next.delete(id);
  else next.add(id);
  return next;
}

/**
 * `none` | `some` | `all` -- what the header box draws. `some` is the
 * half-ticked state, and an empty table is `none` rather than `all`: a box
 * claiming to have selected everything out of nothing reads as a bug.
 */
export function headerState(selected, rows) {
  const ids = selectableIds(rows);
  if (!ids.length || selected.size === 0) return 'none';
  return ids.every((id) => selected.has(id)) ? 'all' : 'some';
}

/** The header box selects everything on screen, or clears it if it already has. */
export const toggleAll = (selected, rows) =>
  (headerState(selected, rows) === 'all' ? new Set() : new Set(selectableIds(rows)));

/**
 * Drops ticks for rows the filters no longer show. Without this, narrowing the
 * search would leave machines ticked off screen and "Remove 3 devices" would
 * delete something nobody could see.
 *
 * Returns the SAME set when there was nothing to drop, so the effect that
 * prunes it does not re-run itself forever on a fresh object.
 */
export function visibleSelection(selected, rows) {
  if (selected.size === 0) return selected;

  const visible = new Set(selectableIds(rows));
  const kept = [...selected].filter((id) => visible.has(id));
  return kept.length === selected.size ? selected : new Set(kept);
}

/** The ticked rows themselves, in the order the table shows them. */
export const selectedDevices = (selected, rows) =>
  rows.filter((row) => isSelectable(row) && selected.has(row.id));

/**
 * The names in a confirm sentence, cut off before it stops fitting on a line.
 * Twenty machine names would push the button that removes them off the side of
 * the bar, which is a poor place to hide the only irreversible control here.
 */
export function describeSelection(devices, limit = 4) {
  const names = devices.map((device) =>
    (device.computerName ? String(device.computerName) : 'an unnamed device'));

  if (names.length <= limit) return names.join(', ');
  return `${names.slice(0, limit).join(', ')} and ${names.length - limit} more`;
}
