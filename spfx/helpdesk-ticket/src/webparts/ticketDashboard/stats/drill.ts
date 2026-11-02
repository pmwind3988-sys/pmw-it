// What a tap on the dashboard lists: the tickets behind one figure, bar,
// segment, square or point. Every list is built from the same per-ticket
// tests the charts count with (stats.ts), and from the same set of tickets the
// chart reads, so the panel's count is always the number that was tapped.

import { QueueTicket } from '../../ticketQueue/queue/ticket';
import {
  AGE_BUCKETS, ageIndex, AREA_LABELS, areaOf, AreaPart, Bucket, closedIn, closedInRange, closerOf, daysToClose,
  GAP_TESTS, GAP_TEXT, heatSlot, HoldState, holdState, holderId, isOpen, openAtEnd, Period, priorityOf, raisedIn,
  raisedInRange, slotText, SPEED_BUCKETS, speedIndex
} from './stats';

export type Drill =
  | { kind: 'open' }
  | { kind: 'raised' }
  | { kind: 'closed' }
  | { kind: 'critical' }
  | { kind: 'bucket'; index: number; series: 'raised' | 'closed' | 'open' }
  | { kind: 'holder'; id: number; state?: HoldState }
  | { kind: 'closer'; id: number }
  | { kind: 'area'; area: string; part?: AreaPart }
  | { kind: 'priority'; priority: string }
  | { kind: 'heat'; day: number; col: number }
  | { kind: 'speed'; index: number }
  | { kind: 'age'; index: number }
  | { kind: 'gap'; index: number };

/** What every list is drawn from: the same sets the dashboard body hands its charts. */
export interface DrillSets {
  /** The filters' people and statuses, any date: the present. */
  scope: QueueTicket[];
  /** The whole filter, dates included. */
  inView: QueueTicket[];
  period: Period;
  buckets: Bucket[];
  now: Date;
  /** "Technician A" for an id, for the panel's heading. */
  nameOf: (id: number) => string;
}

export interface DrillTab {
  label: string;
  drill: Drill;
}

export interface DrillList {
  /** The chart it came from. */
  kicker: string;
  /** What was tapped. */
  title: string;
  /** How the list is ordered, in words. */
  order: string;
  tickets: QueueTicket[];
  /** Raised vs closed offers both halves of the same point. */
  tabs?: DrillTab[];
}

/** What a chart is told: which of its marks is open, if any, and how to open one. */
export interface Picker {
  picked?: string;
  onPick: (d: Drill) => void;
}

export const isPicked = (p: Picker, d: Drill): boolean => p.picked !== undefined && p.picked === drillKey(d);

/** One string per mark: what the chart compares to know which mark is picked. */
export function drillKey(d: Drill): string {
  switch (d.kind) {
    case 'bucket':
      return `bucket:${d.index}:${d.series === 'open' ? 'open' : 'lines'}`;
    case 'holder':
      return `holder:${d.id}:${d.state ?? 'all'}`;
    case 'closer':
      return `closer:${d.id}`;
    case 'area':
      return `area:${d.area}:${d.part ?? 'all'}`;
    case 'priority':
      return `priority:${d.priority}`;
    case 'heat':
      return `heat:${d.day}:${d.col}`;
    case 'speed':
    case 'age':
    case 'gap':
      return `${d.kind}:${d.index}`;
    default:
      return d.kind;
  }
}

const time = (iso: string): number => {
  const n = new Date(iso).getTime();
  return isNaN(n) ? 0 : n;
};
const newestFirst = (a: QueueTicket, b: QueueTicket): number => time(b.created) - time(a.created);
const oldestFirst = (a: QueueTicket, b: QueueTicket): number => time(a.created) - time(b.created);
const closeDays = (t: QueueTicket): number => daysToClose(t) ?? 0;

const NEWEST = 'newest first';

const STATE_WORDS: Record<HoldState, string> = { new: 'New', inProgress: 'In progress', onHold: 'On hold' };
const PART_WORDS: Record<AreaPart, string> = { incident: "Something's broken", request: 'A request', other: 'Other kinds' };

export function drillList(d: Drill, s: DrillSets): DrillList {
  const { scope, inView, period, buckets, now } = s;
  const list = (
    kicker: string,
    title: string,
    pool: QueueTicket[],
    keep: (t: QueueTicket) => boolean,
    order = NEWEST,
    sort: (a: QueueTicket, b: QueueTicket) => number = newestFirst
  ): DrillList => ({ kicker, title, order, tickets: pool.filter(keep).sort(sort) });

  switch (d.kind) {
    case 'open':
      return list('Key figures', 'Open now', inView, isOpen);
    case 'raised':
      return list('Key figures', `Raised ${period.label}`, scope, (t) => raisedInRange(t, period.range));
    case 'closed':
      return list('Key figures', `Closed ${period.label}`, scope, (t) => closedInRange(t, period.range));
    case 'critical':
      return list('Key figures', 'Critical', inView, (t) => t.priority === 'Critical');
    case 'bucket': {
      const b = buckets[d.index];
      if (!b) return { kicker: '', title: '', order: NEWEST, tickets: [] };
      if (d.series === 'open') {
        return list('Still open', `Still open at the end of ${b.label}`, scope, (t) => openAtEnd(t, b.end), 'oldest first', oldestFirst);
      }
      const raised = list('Raised vs closed', b.label, scope, (t) => raisedIn(t, b));
      const closed = list('Raised vs closed', b.label, scope, (t) => closedIn(t, b));
      const tabs: DrillTab[] = [
        { label: `Raised (${raised.tickets.length})`, drill: { kind: 'bucket', index: d.index, series: 'raised' } },
        { label: `Closed (${closed.tickets.length})`, drill: { kind: 'bucket', index: d.index, series: 'closed' } }
      ];
      return { ...(d.series === 'closed' ? closed : raised), tabs };
    }
    case 'holder': {
      const who = d.id === 0 ? 'Unassigned' : s.nameOf(d.id);
      const title = d.state ? `${who} · ${STATE_WORDS[d.state]}` : who;
      return list('Who’s holding what', title, scope, (t) => isOpen(t) && holderId(t) === d.id && (!d.state || holdState(t) === d.state), 'oldest first', oldestFirst);
    }
    case 'closer':
      return list(
        'Closed this month',
        s.nameOf(d.id),
        scope,
        (t) => closerOf(t, now)?.id === d.id,
        'fastest first',
        (a, b) => closeDays(a) - closeDays(b)
      );
    case 'area': {
      const label = d.area ? AREA_LABELS[d.area] ?? d.area : 'No category';
      const title = d.part ? `${label} · ${PART_WORDS[d.part]}` : label;
      return list('What tickets are about', title, inView, (t) => {
        const a = areaOf(t);
        return a.area === d.area && (!d.part || a.part === d.part);
      });
    }
    case 'priority':
      return list('How urgent they were', d.priority, inView, (t) => priorityOf(t) === d.priority);
    case 'heat':
      return list('When tickets arrive', slotText(d.day, d.col), inView, (t) => {
        const slot = heatSlot(t);
        return !!slot && slot.day === d.day && slot.col === d.col;
      });
    case 'speed':
      return list(
        'How long tickets take to close',
        `Time to close: ${SPEED_BUCKETS[d.index].label}`,
        inView,
        (t) => speedIndex(t) === d.index,
        'slowest first',
        (a, b) => closeDays(b) - closeDays(a)
      );
    case 'age':
      return list('How old the open tickets are', `Open ${AGE_BUCKETS[d.index].label.toLowerCase()}`, scope, (t) => ageIndex(t, now) === d.index, 'oldest first', oldestFirst);
    case 'gap': {
      const g = GAP_TESTS[d.index];
      return list('Gaps in the records', GAP_TEXT[d.index].missing, inView, (t) => g.counts(t) && !g.has(t));
    }
    default:
      return { kicker: '', title: '', order: NEWEST, tickets: [] };
  }
}
