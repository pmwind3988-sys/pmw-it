// Pure helpers for the queue screen: how a filter reads in words, how a date
// is shown, how the page is counted, and how saved choices are read back.
// Nothing here touches React, SharePoint or the page.

import { DATE_PRESETS, DateChoice, DatePreset } from '../queue/dateRange';
import { ALL_COLUMN_KEYS, EXPORT_COLUMNS, formatDateTime } from '../queue/exportColumns';
import { defaultFilter, QueueFilter, SortKey, STATUSES, UNASSIGNED } from '../queue/queueFilter';
import { PersonRef, QueueTicket } from '../queue/ticket';

export const PAGE_SIZE = 50;
export const OPEN_STATUSES = ['New', 'In progress', 'On Hold'];

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const pad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

/** "29 Oct 2026, 2:05 PM", local time. Empty when the value is not a date. */
export function formatWhen(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const hour24 = d.getHours();
  const hour12 = hour24 % 12 === 0 ? 12 : hour24 % 12;
  const part = hour24 < 12 ? 'AM' : 'PM';
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}, ${hour12}:${pad2(d.getMinutes())} ${part}`;
}

/** 'YYYY-MM' for a month input. */
export function monthKey(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}

/** 'YYYY-MM-DD' for a date input. */
export function dayKey(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

export function countText(n: number): string {
  return n === 1 ? '1 ticket' : `${n} tickets`;
}

/** Who a ticket was raised by, and the department, on one line under the title. */
export function raisedLine(t: QueueTicket): string {
  const who = t.raisedBy?.name ? `Raised by ${t.raisedBy.name}` : '';
  return [who, t.department].filter((s) => s !== '').join(' · ');
}

export function sameSet(a: string[], b: string[]): boolean {
  return a.length === b.length && a.every((x) => b.indexOf(x) !== -1);
}

/**
 * The people a set of assignee ids names, in the order the screen lists them:
 * me first, Unassigned second, then everyone else A–Z.
 */
export function assigneeNames(
  ids: number[],
  me: PersonRef,
  nameOf: (id: number) => string,
  meWord: string
): string[] {
  const rank = (id: number): number => (id === me.id ? 0 : id === UNASSIGNED ? 1 : 2);
  return ids
    .slice()
    .sort((a, b) => rank(a) - rank(b) || nameOf(a).localeCompare(nameOf(b)))
    .map((id) => (id === me.id ? meWord : id === UNASSIGNED ? 'Unassigned' : nameOf(id)));
}

/** The button's words: "Everyone", "Me", "Me + 2 others", "Papathy, Ashraf". */
export function assigneeSummary(ids: number[], me: PersonRef, nameOf: (id: number) => string): string {
  if (!ids.length) return 'Everyone';
  const names = assigneeNames(ids, me, nameOf, 'Me');
  if (names.length === 1) return names[0];
  const others = names.length - 1;
  return `${names[0]} + ${others} other${others === 1 ? '' : 's'}`;
}

const SORT_WORDS: Record<SortKey, { label: string; asc: string; desc: string }> = {
  number: { label: 'Ticket ID', asc: 'A–Z', desc: 'Z–A' },
  priority: { label: 'Priority', asc: 'Critical first', desc: 'Low first' },
  status: { label: 'Status', asc: 'New first', desc: 'Cancelled first' },
  assignedTo: { label: 'Assigned to', asc: 'A–Z', desc: 'Z–A' },
  created: { label: 'Raised on', asc: 'oldest first', desc: 'newest first' },
  modified: { label: 'Updated', asc: 'oldest first', desc: 'newest first' }
};

/** "Raised on (newest first)" for the export's "Sorted by" line. */
export function sortDescription(key: SortKey, desc: boolean): string {
  const w = SORT_WORDS[key];
  return `${w.label} (${desc ? w.desc : w.asc})`;
}

/** Dates start newest first; every other column starts A–Z or most-important first. */
export function initialDesc(key: SortKey): boolean {
  return key === 'created' || key === 'modified';
}

export function pageCount(total: number): number {
  return Math.max(1, Math.ceil(total / PAGE_SIZE));
}

export interface ExportContext {
  exportedBy: string;
  assignees: string;
  raised: string;
  statuses: string[];
  search: string;
  sortText: string;
  fileParts: string[];
}

/** The "About this export" sheet, as label/value pairs. */
export function exportLines(c: ExportContext, count: number, now: Date): [string, string][] {
  return [
    ['Exported by', c.exportedBy || '—'],
    ['Exported on', formatDateTime(now)],
    ['Assigned to', c.assignees],
    ['Raised', c.raised],
    ['Status', c.statuses.length ? c.statuses.join(', ') : 'Any'],
    ['Search', c.search.trim() || '—'],
    ['Sorted by', c.sortText],
    ['Tickets', String(count)]
  ];
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

function parseJson(raw: string): unknown {
  try {
    return JSON.parse(raw) as unknown;
  } catch {
    return undefined;
  }
}

const text = (v: unknown): string | undefined => (typeof v === 'string' ? v : undefined);

/**
 * The saved filter, checked field by field. Anything unreadable falls back to
 * the default for this person; a bad field falls back on its own.
 */
export function parseFilter(raw: string | undefined, myId: number): QueueFilter {
  const fallback = defaultFilter(myId);
  if (raw === undefined) return fallback;
  const parsed = parseJson(raw);
  if (!isRecord(parsed)) return fallback;

  const list: unknown = parsed.assignees;
  const assignees = Array.isArray(list)
    ? list.filter((x: unknown): x is number => typeof x === 'number' && isFinite(x))
    : fallback.assignees;

  const dateRaw: unknown = parsed.date;
  const date: Record<string, unknown> = isRecord(dateRaw) ? dateRaw : {};
  const presetRaw = date.preset;
  const preset: DatePreset = DATE_PRESETS.some((p) => p.id === presetRaw) ? (presetRaw as DatePreset) : 'all';
  const choice: DateChoice = { preset, month: text(date.month), start: text(date.start), end: text(date.end) };

  const statusList: unknown = parsed.statuses;
  const picked = Array.isArray(statusList)
    ? statusList.filter((s: unknown): s is string => typeof s === 'string')
    : [];
  const statuses = STATUSES.filter((s) => picked.indexOf(s) !== -1);

  return { assignees, date: choice, statuses, search: text(parsed.search) ?? '' };
}

/** The saved export columns: unknown keys dropped; nothing usable means all of them. */
export function parseColumns(raw: string | undefined): string[] {
  if (raw === undefined) return ALL_COLUMN_KEYS.slice();
  const parsed = parseJson(raw);
  if (!Array.isArray(parsed)) return ALL_COLUMN_KEYS.slice();
  const picked = parsed.filter((k: unknown): k is string => typeof k === 'string' && ALL_COLUMN_KEYS.indexOf(k) !== -1);
  if (!picked.length) return ALL_COLUMN_KEYS.slice();
  return EXPORT_COLUMNS.map((c) => c.key).filter((k) => picked.indexOf(k) !== -1);
}
