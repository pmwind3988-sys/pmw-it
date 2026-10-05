import { labelFor } from '../fieldGroups.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';

/** These belong to the owner history; showing them twice would make one move look like two. */
const OWNER_FIELDS = new Set(['owner', 'department', 'location', 'status', 'device']);

export function specHistory(changes = []) {
  const byDay = new Map();
  const sorted = [...changes]
    .filter((change) => !OWNER_FIELDS.has(change.fieldName) && typeof change.changedOn === 'number')
    .sort((a, b) => b.changedOn - a.changedOn);

  for (const change of sorted) {
    const dayLabel = formatMYT(change.changedOn, 'date');
    if (!byDay.has(dayLabel)) byDay.set(dayLabel, { day: change.changedOn, dayLabel, rows: [] });
    byDay.get(dayLabel).rows.push({
      fieldName: change.fieldName,
      label: change.fieldName === 'computerName' ? 'Renamed' : labelFor(change.fieldName),
      oldValue: change.oldValue,
      newValue: change.newValue,
      rename: change.fieldName === 'computerName',
    });
  }

  // Oldest-first within a day reads as the order things happened.
  return [...byDay.values()].map((group) => ({ ...group, rows: group.rows.reverse() }));
}
