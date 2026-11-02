// Which columns the queue's table shows, and in what order. Each owner picks
// their own in this browser; the queue as it always was, plus "Time to
// complete", is the default. Ticket ID is the link to the ticket, so it is
// always there and always first.

import { SortKey } from './queueFilter';

export interface TableColumn {
  key: string;
  label: string;
  sort?: SortKey;
}

export const TABLE_COLUMNS: TableColumn[] = [
  { key: 'number', label: 'Ticket ID', sort: 'number' },
  { key: 'title', label: 'Short description' },
  { key: 'category', label: 'Category' },
  { key: 'priority', label: 'Priority', sort: 'priority' },
  { key: 'status', label: 'Status', sort: 'status' },
  { key: 'assignedTo', label: 'Assigned to', sort: 'assignedTo' },
  { key: 'created', label: 'Raised on', sort: 'created' },
  { key: 'modified', label: 'Updated', sort: 'modified' },
  { key: 'timeToComplete', label: 'Time to complete', sort: 'timeToComplete' },
  { key: 'completedDate', label: 'Completed on', sort: 'completedDate' },
  { key: 'completedBy', label: 'Completed by' },
  { key: 'raisedBy', label: 'Raised by' },
  { key: 'raisedByEmail', label: 'Raised by (email)' },
  { key: 'department', label: 'Department' },
  { key: 'description', label: 'Issue description' },
  { key: 'cc', label: 'Cc' },
  { key: 'rootCause', label: 'Root cause' },
  { key: 'resolution', label: 'Resolution summary' },
  { key: 'comments', label: 'Comments' },
  { key: 'holdCount', label: 'Times on hold' },
  { key: 'totalHoldDays', label: 'Days on hold' },
  { key: 'kbCandidate', label: 'KB candidate' }
];

/** Always shown, always first. */
export const LOCKED_COLUMN = 'number';

export const DEFAULT_COLUMNS = [
  'number', 'title', 'category', 'priority', 'status', 'assignedTo', 'created', 'modified', 'timeToComplete'
];

const KNOWN = TABLE_COLUMNS.map((c) => c.key);

export const columnOf = (key: string): TableColumn | undefined => TABLE_COLUMNS.filter((c) => c.key === key)[0];

/** Ticket ID first, every key known, none twice. */
function tidy(keys: string[]): string[] {
  const out: string[] = [LOCKED_COLUMN];
  keys.forEach((k) => {
    if (KNOWN.indexOf(k) !== -1 && out.indexOf(k) === -1) out.push(k);
  });
  return out;
}

/** The saved choice, checked: unknown keys dropped; nothing usable means the default. */
export function parseTableColumns(raw: string | undefined): string[] {
  if (raw === undefined) return DEFAULT_COLUMNS.slice();
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return DEFAULT_COLUMNS.slice();
  }
  if (!Array.isArray(parsed)) return DEFAULT_COLUMNS.slice();
  const keys = tidy(parsed.filter((k: unknown): k is string => typeof k === 'string'));
  return keys.length > 1 ? keys : DEFAULT_COLUMNS.slice();
}

/** The column at `from` moved to sit at `to`; Ticket ID never moves and nothing goes before it. */
export function moveColumn(keys: string[], from: number, to: number): string[] {
  if (from <= 0 || from >= keys.length) return keys;
  const target = Math.max(1, Math.min(keys.length - 1, to));
  if (target === from) return keys;
  const out = keys.slice();
  const [k] = out.splice(from, 1);
  out.splice(target, 0, k);
  return out;
}

export function removeColumn(keys: string[], key: string): string[] {
  return key === LOCKED_COLUMN ? keys : keys.filter((k) => k !== key);
}

/** A column added from the list goes on the end. */
export function addColumn(keys: string[], key: string): string[] {
  return tidy(keys.concat([key]));
}

/** The columns not on show, in the list's own order. */
export function hiddenColumns(keys: string[]): TableColumn[] {
  return TABLE_COLUMNS.filter((c) => keys.indexOf(c.key) === -1);
}

export function isDefault(keys: string[]): boolean {
  return keys.length === DEFAULT_COLUMNS.length && keys.every((k, i) => DEFAULT_COLUMNS[i] === k);
}
