// Which tickets the queue shows: who they are assigned to, when they were
// raised, their status, and what was typed in the search box.

import { DateChoice, inRange, rangeOf } from './dateRange';
import { PersonRef, QueueTicket } from './ticket';

/** The assignee filter's entry for tickets nobody has been given yet. */
export const UNASSIGNED = 0;

export interface QueueFilter {
  /** Site user ids; UNASSIGNED stands for "nobody". Empty means everyone. */
  assignees: number[];
  date: DateChoice;
  /** Empty means every status. */
  statuses: string[];
  search: string;
}

export const STATUSES = ['New', 'In progress', 'On Hold', 'Completed', 'Cancelled'];

/** What an owner sees first: their own open tickets, any time. */
export function defaultFilter(myId: number): QueueFilter {
  return { assignees: [myId], date: { preset: 'all' }, statuses: ['New', 'In progress', 'On Hold'], search: '' };
}

function haystack(t: QueueTicket): string {
  return [
    t.number, t.title, t.description, t.category, t.department,
    t.raisedBy?.name, t.raisedBy?.email, t.assignedTo?.name
  ].join(' ').toLowerCase();
}

export function applyFilter(rows: QueueTicket[], f: QueueFilter, now: Date = new Date()): QueueTicket[] {
  const range = rangeOf(f.date, now);
  const needles = f.search.toLowerCase().split(/\s+/).filter(Boolean);
  return rows.filter((t) => {
    if (f.assignees.length) {
      const who = t.assignedTo ? t.assignedTo.id : UNASSIGNED;
      if (f.assignees.indexOf(who) === -1) return false;
    }
    if (f.statuses.length && f.statuses.indexOf(t.status) === -1) return false;
    if (!inRange(t.created, range)) return false;
    if (needles.length) {
      const hay = haystack(t);
      // Every word typed must appear somewhere: "sap ipoh" narrows, never widens.
      if (!needles.every((n) => hay.indexOf(n) !== -1)) return false;
    }
    return true;
  });
}

/** Everyone tickets have been assigned to, A–Z, with how many each holds. */
export function assigneesOf(rows: QueueTicket[]): { person: PersonRef; count: number }[] {
  const byId = new Map<number, { person: PersonRef; count: number }>();
  rows.forEach((t) => {
    if (!t.assignedTo) return;
    const hit = byId.get(t.assignedTo.id);
    if (hit) hit.count += 1;
    else byId.set(t.assignedTo.id, { person: t.assignedTo, count: 1 });
  });
  const list: { person: PersonRef; count: number }[] = [];
  byId.forEach((v) => list.push(v));
  return list.sort((a, b) => a.person.name.localeCompare(b.person.name));
}

export type SortKey = 'created' | 'modified' | 'number' | 'priority' | 'status' | 'assignedTo';

const PRIORITY_RANK: Record<string, number> = { Critical: 0, High: 1, Normal: 2, Low: 3, Undefined: 4 };
const STATUS_RANK: Record<string, number> = { New: 0, 'In progress': 1, 'On Hold': 2, Completed: 3, Cancelled: 4 };

/** Sorted copy. Ties fall back to newest first, so the order never shuffles between renders. */
export function sortTickets(rows: QueueTicket[], key: SortKey, desc: boolean): QueueTicket[] {
  const value = (t: QueueTicket): string | number => {
    switch (key) {
      case 'created': return t.created;
      case 'modified': return t.modified;
      case 'number': return t.number;
      case 'priority': return PRIORITY_RANK[t.priority] ?? 9;
      case 'status': return STATUS_RANK[t.status] ?? 9;
      case 'assignedTo': return t.assignedTo ? t.assignedTo.name.toLowerCase() : '￿';
    }
  };
  return rows.slice().sort((a, b) => {
    const va = value(a);
    const vb = value(b);
    const c = va < vb ? -1 : va > vb ? 1 : 0;
    if (c !== 0) return desc ? -c : c;
    return a.created < b.created ? 1 : a.created > b.created ? -1 : b.id - a.id;
  });
}
