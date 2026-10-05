import { formatItems } from './toChecklistItem.js';

/**
 * A checklist's values as words, for the places that show them rather than
 * edit them: a field IT locked on a shared link, and the signed copy.
 */

export const FIELD_LABELS = {
  employeeName: 'Employee Name',
  employeeNo: 'Employee No',
  position: 'Position',
  entity: 'Entity',
  formDate: 'Date',
  checkedItems: 'Asset Checklist',
  items: 'Requested Items',
  serialNumbers: 'Serial Numbers',
  otherRemarks: 'Other Remarks',
};

/** `2026-10-05` → `5 October 2026`, read as the calendar day it names. */
export function dayLabel(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value ?? ''));
  if (!match) return String(value ?? '');
  const [, year, month, day] = match.map(Number);
  return new Intl.DateTimeFormat('en-GB', {
    day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC',
  }).format(Date.UTC(year, month - 1, day));
}

/** Item lines as the list stores them: `Laptop x 1`, one per line. */
export function itemLines(rows) {
  return formatItems(rows).split('\n').filter(Boolean);
}

/** One value, as text. Lists come back as arrays of lines. */
export function describeValue(field, value) {
  if (field === 'formDate') return dayLabel(value);
  if (field === 'items') return itemLines(value);
  if (field === 'checkedItems') return (value ?? []).filter(Boolean);
  return String(value ?? '').trim();
}
