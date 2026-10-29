// What can go into an export, and the CSV writer. The .xlsx writer (xlsx.ts)
// reads the same columns, so the two files can never disagree on content.

import { QueueTicket, PersonRef } from './ticket';

export type CellValue = string | number | Date | boolean | undefined;

export interface ExportColumn {
  key: string;
  header: string;
  /** Excel column width, in characters. */
  width: number;
  /** Long text: wrap it in the spreadsheet. */
  wrap?: boolean;
  read: (t: QueueTicket) => CellValue;
}

const who = (p?: PersonRef): string => (p ? p.name : '');
const when = (iso: string): Date | undefined => {
  if (!iso) return undefined;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : d;
};

export const EXPORT_COLUMNS: ExportColumn[] = [
  { key: 'number', header: 'Ticket ID', width: 15, read: (t) => t.number },
  { key: 'title', header: 'Short description', width: 40, wrap: true, read: (t) => t.title },
  { key: 'description', header: 'Issue description', width: 60, wrap: true, read: (t) => t.description },
  { key: 'category', header: 'Category', width: 28, read: (t) => t.category },
  { key: 'priority', header: 'Priority', width: 10, read: (t) => t.priority },
  { key: 'status', header: 'Status', width: 12, read: (t) => t.status },
  { key: 'assignedTo', header: 'Assigned to', width: 26, read: (t) => who(t.assignedTo) },
  { key: 'raisedBy', header: 'Raised by', width: 26, read: (t) => who(t.raisedBy) },
  { key: 'raisedByEmail', header: 'Raised by (email)', width: 30, read: (t) => t.raisedBy?.email ?? '' },
  { key: 'department', header: 'Department', width: 16, read: (t) => t.department },
  { key: 'cc', header: 'Cc', width: 30, wrap: true, read: (t) => t.cc.map((p) => p.name).join('; ') },
  { key: 'created', header: 'Raised on', width: 18, read: (t) => when(t.created) },
  { key: 'modified', header: 'Last updated', width: 18, read: (t) => when(t.modified) },
  { key: 'completedDate', header: 'Completed on', width: 18, read: (t) => when(t.completedDate) },
  { key: 'completedBy', header: 'Completed by', width: 26, read: (t) => who(t.completedBy) },
  { key: 'rootCause', header: 'Root cause', width: 20, read: (t) => t.rootCause },
  { key: 'resolution', header: 'Resolution summary', width: 50, wrap: true, read: (t) => t.resolution },
  { key: 'comments', header: 'Comments', width: 50, wrap: true, read: (t) => t.comments },
  { key: 'holdCount', header: 'Times on hold', width: 13, read: (t) => t.holdCount },
  { key: 'totalHoldDays', header: 'Days on hold', width: 13, read: (t) => t.totalHoldDays },
  { key: 'kbCandidate', header: 'KB candidate', width: 13, read: (t) => (t.kbCandidate ? 'Yes' : 'No') }
];

export const ALL_COLUMN_KEYS = EXPORT_COLUMNS.map((c) => c.key);

/** The chosen columns, always in the sheet's own order whatever order they were ticked in. */
export function chosenColumns(keys: string[]): ExportColumn[] {
  return EXPORT_COLUMNS.filter((c) => keys.indexOf(c.key) !== -1);
}

const pad = (n: number): string => (n < 10 ? `0${n}` : String(n));

/** Local date and time, the way Excel reads it back as a date: 2026-10-29 14:05. */
export function formatDateTime(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/**
 * A cell Excel would run as a formula ("=HYPERLINK(...)", "+60 12...", "@SUM")
 * is a hazard in a file built from what staff typed. Prefix it with an
 * apostrophe so it stays text. A phone number loses nothing by it.
 */
export function safeText(s: string): string {
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
}

function csvCell(v: CellValue): string {
  let s: string;
  if (v === undefined) s = '';
  else if (v instanceof Date) s = formatDateTime(v);
  else if (typeof v === 'boolean') s = v ? 'Yes' : 'No';
  else if (typeof v === 'number') s = String(v);
  else s = safeText(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

/**
 * CSV with a byte-order mark, so Excel opens it as UTF-8 and a name like
 * "Muñoz" or a Chinese name survives, and CRLF line ends, which Excel expects.
 */
export function toCsv(rows: QueueTicket[], columns: ExportColumn[]): string {
  const lines = [columns.map((c) => csvCell(c.header)).join(',')];
  rows.forEach((t) => lines.push(columns.map((c) => csvCell(c.read(t))).join(',')));
  return '﻿' + lines.join('\r\n') + '\r\n';
}

/** "IT tickets - Papathy, Ashraf - Oct 2026 - 2026-10-29.xlsx", safe for every file system. */
export function exportFileName(parts: string[], ext: 'csv' | 'xlsx', now: Date = new Date()): string {
  const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  const base = ['IT tickets', ...parts.filter(Boolean), stamp]
    .join(' - ')
    .replace(/[\\/:*?"<>|]+/g, '')
    .replace(/\s+/g, ' ')
    .slice(0, 150)
    .trim();
  return `${base}.${ext}`;
}
