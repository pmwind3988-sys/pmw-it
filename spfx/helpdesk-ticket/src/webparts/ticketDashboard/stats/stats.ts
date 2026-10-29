// Every figure on the IT helpdesk dashboard, worked out from plain ticket rows.
//
// Two sets of tickets feed it, and which figure reads which is the whole
// design:
//   - `inView`: the queue's filter as set, dates included. Distributions
//     (what tickets are about, how urgent, how long they took, when they
//     arrive) describe exactly the tickets the filters picked, and so does Export.
//   - `scope`: the same filter WITHOUT its dates. "Open now", "this month vs
//     last month", who holds what and how old the open tickets are describe the
//     present, and a date filter of "March" must not make 41 open tickets read 3.
// The trend charts read `scope` too, cut to the chosen dates (or the past 12
// months when none are chosen), because a trend needs time on its x axis.

import { Range } from '../../ticketQueue/queue/dateRange';
import { QueueTicket } from '../../ticketQueue/queue/ticket';

const DAY = 24 * 60 * 60 * 1000;
const CLOSED = ['Completed', 'Cancelled'];

export const isOpen = (t: QueueTicket): boolean => CLOSED.indexOf(t.status) === -1;

/** When a ticket stopped being open: its completed date, else (closed without one) its last update. */
export function closedAt(t: QueueTicket): Date | undefined {
  if (isOpen(t)) return undefined;
  const iso = t.completedDate || t.modified;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : d;
}

/** Days from raised to completed, for Completed tickets only (a cancellation is not a fix). */
export function daysToClose(t: QueueTicket): number | undefined {
  if (t.status !== 'Completed') return undefined;
  const end = closedAt(t);
  const start = new Date(t.created);
  if (!end || isNaN(start.getTime())) return undefined;
  return Math.max(0, (end.getTime() - start.getTime()) / DAY);
}

export function median(xs: number[]): number | undefined {
  if (!xs.length) return undefined;
  const s = xs.slice().sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** The value below which `p` of the list falls (nearest rank). */
export function percentile(xs: number[], p: number): number | undefined {
  if (!xs.length) return undefined;
  const s = xs.slice().sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.max(0, Math.ceil(p * s.length) - 1))];
}

// ---------------------------------------------------------------- trend

export type Grain = 'day' | 'week' | 'month';

export interface Bucket {
  /** Start of the bucket, local. */
  start: Date;
  end: Date;
  label: string;
  raised: number;
  closed: number;
  /** Tickets still open when the bucket ended. */
  openAtEnd: number;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const startOfDay = (d: Date): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate());

/** A day per point up to a month, a week per point up to four months, otherwise a month per point. */
export function grainFor(from: Date, to: Date): Grain {
  const days = (to.getTime() - from.getTime()) / DAY;
  if (days <= 31) return 'day';
  // Four calendar months run to 123 days.
  if (days <= 124) return 'week';
  return 'month';
}

/**
 * The window the trend charts cover: the chosen dates, or the past twelve
 * calendar months (this one included) when "Any time" is chosen. A window that
 * runs into the future stops at today, so a "This month" line does not fall
 * to zero for the days that have not happened.
 */
export function trendWindow(range: Range, now: Date): { from: Date; to: Date } {
  const today = startOfDay(now);
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  const from = range.from ?? new Date(now.getFullYear(), now.getMonth() - 11, 1);
  let to = range.to ?? tomorrow;
  if (to > tomorrow) to = tomorrow;
  return { from, to: to > from ? to : tomorrow };
}

function bucketStarts(from: Date, to: Date, grain: Grain): Date[] {
  const out: Date[] = [];
  let d: Date;
  if (grain === 'month') d = new Date(from.getFullYear(), from.getMonth(), 1);
  else if (grain === 'week') {
    const s = startOfDay(from);
    d = new Date(s.getFullYear(), s.getMonth(), s.getDate() - ((s.getDay() + 6) % 7));
  } else d = startOfDay(from);
  while (d < to && out.length < 400) {
    out.push(d);
    d =
      grain === 'month'
        ? new Date(d.getFullYear(), d.getMonth() + 1, 1)
        : new Date(d.getFullYear(), d.getMonth(), d.getDate() + (grain === 'week' ? 7 : 1));
  }
  return out;
}

function labelOf(d: Date, grain: Grain, multiYear: boolean): string {
  if (grain === 'month') return multiYear || d.getMonth() === 0 ? `${MONTHS[d.getMonth()]} ${String(d.getFullYear()).slice(2)}` : MONTHS[d.getMonth()];
  return `${d.getDate()} ${MONTHS[d.getMonth()]}`;
}

export function trend(scope: QueueTicket[], range: Range, now: Date): { grain: Grain; buckets: Bucket[] } {
  const win = trendWindow(range, now);
  const grain = grainFor(win.from, win.to);
  const starts = bucketStarts(win.from, win.to, grain);
  const multiYear = starts.length > 0 && starts[0].getFullYear() !== starts[starts.length - 1].getFullYear();
  const buckets: Bucket[] = starts.map((start, i) => ({
    start,
    end: i + 1 < starts.length ? starts[i + 1] : win.to,
    label: labelOf(start, grain, multiYear),
    raised: 0,
    closed: 0,
    openAtEnd: 0
  }));
  const find = (d: Date): Bucket | undefined => buckets.find((b) => d >= b.start && d < b.end);
  scope.forEach((t) => {
    const made = new Date(t.created);
    const shut = closedAt(t);
    const rb = isNaN(made.getTime()) ? undefined : find(made);
    if (rb) rb.raised += 1;
    if (shut && t.status === 'Completed') {
      const cb = find(shut);
      if (cb) cb.closed += 1;
    }
    buckets.forEach((b) => {
      if (made < b.end && (!shut || shut >= b.end)) b.openAtEnd += 1;
    });
  });
  return { grain, buckets };
}

// ---------------------------------------------------------------- headline figures

export interface Headline {
  openNow: number;
  unassignedOpen: number;
  raisedThisMonth: number;
  raisedLastMonth: number;
  closedThisMonth: number;
  closedLastMonth: number;
  /** Median days to close, tickets in view. */
  medianDays?: number;
  /** 90th percentile days to close, tickets in view. */
  p90Days?: number;
  oldestOpen?: { ticket: QueueTicket; days: number };
  criticalInView: number;
  inView: number;
}

const monthStart = (now: Date, back: number): Date => new Date(now.getFullYear(), now.getMonth() - back, 1);

export function headline(scope: QueueTicket[], inView: QueueTicket[], now: Date): Headline {
  const thisM = monthStart(now, 0);
  const nextM = monthStart(now, -1);
  const lastM = monthStart(now, 1);
  const between = (d: Date | undefined, a: Date, b: Date): boolean => !!d && d >= a && d < b;
  const open = scope.filter(isOpen);
  let oldest: { ticket: QueueTicket; days: number } | undefined;
  open.forEach((t) => {
    const days = Math.floor((now.getTime() - new Date(t.created).getTime()) / DAY);
    if (!isNaN(days) && (!oldest || days > oldest.days)) oldest = { ticket: t, days };
  });
  const done = scope.filter((t) => t.status === 'Completed');
  const speeds = inView.map(daysToClose).filter((d): d is number => d !== undefined);
  return {
    openNow: open.length,
    unassignedOpen: open.filter((t) => !t.assignedTo).length,
    raisedThisMonth: scope.filter((t) => between(new Date(t.created), thisM, nextM)).length,
    raisedLastMonth: scope.filter((t) => between(new Date(t.created), lastM, thisM)).length,
    closedThisMonth: done.filter((t) => between(closedAt(t), thisM, nextM)).length,
    closedLastMonth: done.filter((t) => between(closedAt(t), lastM, thisM)).length,
    medianDays: median(speeds),
    p90Days: percentile(speeds, 0.9),
    oldestOpen: oldest,
    criticalInView: inView.filter((t) => t.priority === 'Critical').length,
    inView: inView.length
  };
}

/** "↑ 29% on last month", "↓ 12% on last month", "Same as last month", "" when last month had none. */
export function changeText(now: number, before: number): string {
  if (before === 0) return now === 0 ? 'Same as last month' : '';
  const pct = Math.round(((now - before) / before) * 100);
  if (pct === 0) return 'Same as last month';
  return `${pct > 0 ? '↑' : '↓'} ${Math.abs(pct)}% on last month`;
}

// ---------------------------------------------------------------- people

export interface Holder {
  id: number;
  name: string;
  new: number;
  inProgress: number;
  onHold: number;
  total: number;
}

/** Open tickets per person (and "Unassigned"), most first. */
export function holders(scope: QueueTicket[]): Holder[] {
  const map = new Map<number, Holder>();
  scope.filter(isOpen).forEach((t) => {
    const id = t.assignedTo ? t.assignedTo.id : 0;
    let h = map.get(id);
    if (!h) {
      h = { id, name: t.assignedTo ? t.assignedTo.name : 'Unassigned', new: 0, inProgress: 0, onHold: 0, total: 0 };
      map.set(id, h);
    }
    if (t.status === 'In progress') h.inProgress += 1;
    else if (t.status === 'On Hold') h.onHold += 1;
    else h.new += 1;
    h.total += 1;
  });
  const list: Holder[] = [];
  map.forEach((h) => list.push(h));
  return list.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
}

export interface Closer {
  id: number;
  name: string;
  closed: number;
  medianDays?: number;
}

/**
 * Who closed what this calendar month. Credited to "Completed by" where the
 * list recorded it, else to whoever the ticket is assigned to.
 */
export function leaderboard(scope: QueueTicket[], now: Date): { closers: Closer[]; fastest?: Closer } {
  const from = monthStart(now, 0);
  const to = monthStart(now, -1);
  const map = new Map<number, { c: Closer; speeds: number[] }>();
  scope.forEach((t) => {
    if (t.status !== 'Completed') return;
    const at = closedAt(t);
    if (!at || at < from || at >= to) return;
    const who = t.completedBy ?? t.assignedTo;
    if (!who) return;
    let e = map.get(who.id);
    if (!e) {
      e = { c: { id: who.id, name: who.name, closed: 0 }, speeds: [] };
      map.set(who.id, e);
    }
    e.c.closed += 1;
    const d = daysToClose(t);
    if (d !== undefined) e.speeds.push(d);
  });
  const closers: Closer[] = [];
  map.forEach((e) => closers.push({ ...e.c, medianDays: median(e.speeds) }));
  closers.sort((a, b) => b.closed - a.closed || a.name.localeCompare(b.name));
  // "Fastest" only counts somebody with at least three closes: one quick ticket is luck.
  let fastest: Closer | undefined;
  closers.forEach((c) => {
    if (c.closed >= 3 && c.medianDays !== undefined && (!fastest || c.medianDays < (fastest.medianDays as number))) fastest = c;
  });
  return { closers, fastest };
}

// ---------------------------------------------------------------- what and how urgent

/** Old names the Category column used to hold, read as the area they belong to now. */
const AREA_ALIASES: Record<string, string> = {
  Software: 'Application/Software',
  Application: 'Application/Software',
  'Access/Security': 'Access'
};

export const AREA_LABELS: Record<string, string> = {
  Hardware: 'Laptop, PC or printer',
  SAP: 'SAP',
  'Application/Software': 'Software & apps',
  Email: 'Email & Teams',
  Network: 'Internet & Wi-Fi',
  Access: 'Access & passwords',
  Security: 'Security',
  'How-To/Training': 'How do I…?',
  'Onboarding/Offboarding': 'Joining or leaving'
};

export interface AreaCount {
  area: string;
  label: string;
  incident: number;
  request: number;
  /** Change, or no kind recorded. */
  other: number;
  total: number;
}

/** "SAP - Incident" → { area: 'SAP', kind: 'Incident' }; old names folded into their area. */
export function splitCategory(category: string): { area: string; kind: string } {
  const c = category.trim();
  if (!c) return { area: '', kind: '' };
  const i = c.lastIndexOf(' - ');
  const rawArea = i === -1 ? c : c.slice(0, i).trim();
  const kind = i === -1 ? '' : c.slice(i + 3).trim();
  return { area: AREA_ALIASES[rawArea] ?? rawArea, kind };
}

export function byArea(inView: QueueTicket[]): { areas: AreaCount[]; uncategorised: number } {
  const map = new Map<string, AreaCount>();
  let uncategorised = 0;
  inView.forEach((t) => {
    const { area, kind } = splitCategory(t.category);
    if (!area) {
      uncategorised += 1;
      return;
    }
    let a = map.get(area);
    if (!a) {
      a = { area, label: AREA_LABELS[area] ?? area, incident: 0, request: 0, other: 0, total: 0 };
      map.set(area, a);
    }
    if (kind === 'Incident') a.incident += 1;
    else if (kind === 'Request') a.request += 1;
    else a.other += 1;
    a.total += 1;
  });
  const areas: AreaCount[] = [];
  map.forEach((a) => areas.push(a));
  return { areas: areas.sort((a, b) => b.total - a.total || a.label.localeCompare(b.label)), uncategorised };
}

export const PRIORITY_ORDER = ['Critical', 'High', 'Normal', 'Low'];

export function byPriority(inView: QueueTicket[]): { priority: string; count: number }[] {
  const counts: Record<string, number> = {};
  inView.forEach((t) => {
    const p = PRIORITY_ORDER.indexOf(t.priority) === -1 ? 'Not set' : t.priority;
    counts[p] = (counts[p] ?? 0) + 1;
  });
  return PRIORITY_ORDER.concat(['Not set'])
    .filter((p) => counts[p])
    .map((p) => ({ priority: p, count: counts[p] }));
}

// ---------------------------------------------------------------- when and how long

export const HEAT_DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
/** Columns: before 8 AM, each hour 8 AM–5 PM, 6 PM and after. */
export const HEAT_HOURS = ['Early', '8', '9', '10', '11', '12', '1', '2', '3', '4', '5', 'Late'];

export interface Heat {
  /** [day 0=Mon][column], counts. */
  grid: number[][];
  max: number;
  busiest?: { day: number; col: number; count: number };
  /** Saturday and Sunday rows are dropped when both are empty. */
  showWeekend: boolean;
}

export function heatColumn(hour: number): number {
  if (hour < 8) return 0;
  if (hour >= 18) return 11;
  return hour - 7;
}

export function arrivals(inView: QueueTicket[]): Heat {
  const grid = HEAT_DAYS.map(() => HEAT_HOURS.map(() => 0));
  inView.forEach((t) => {
    const d = new Date(t.created);
    if (isNaN(d.getTime())) return;
    grid[(d.getDay() + 6) % 7][heatColumn(d.getHours())] += 1;
  });
  let max = 0;
  let busiest: Heat['busiest'];
  grid.forEach((row, day) =>
    row.forEach((count, col) => {
      if (count > max) {
        max = count;
        busiest = { day, col, count };
      }
    })
  );
  const weekend = grid[5].concat(grid[6]).some((n) => n > 0);
  return { grid, max, busiest, showWeekend: weekend };
}

/** "Thursday, 3–4 PM", "Monday, before 8 AM". */
export function slotText(day: number, col: number): string {
  const names = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  if (col === 0) return `${names[day]}, before 8 AM`;
  if (col === 11) return `${names[day]}, after 6 PM`;
  const h = col + 7;
  const fmt = (x: number): string => `${x > 12 ? x - 12 : x}`;
  return `${names[day]}, ${fmt(h)}–${fmt(h + 1)} ${h + 1 >= 12 ? 'PM' : 'AM'}`;
}

export const SPEED_BUCKETS = [
  { label: 'Same day', max: 1 },
  { label: '1–2 days', max: 3 },
  { label: '3–7 days', max: 8 },
  { label: '1–4 weeks', max: 29 },
  { label: 'Over a month', max: Infinity }
];

export function speeds(inView: QueueTicket[]): { label: string; count: number }[] {
  const counts = SPEED_BUCKETS.map(() => 0);
  inView.forEach((t) => {
    const d = daysToClose(t);
    if (d === undefined) return;
    counts[SPEED_BUCKETS.findIndex((b) => d < b.max)] += 1;
  });
  return SPEED_BUCKETS.map((b, i) => ({ label: b.label, count: counts[i] }));
}

export const AGE_BUCKETS = [
  { label: 'Under a week', max: 7 },
  { label: '1–4 weeks', max: 28 },
  { label: '1–3 months', max: 91 },
  { label: '3–6 months', max: 183 },
  { label: 'Over 6 months', max: Infinity }
];

export function ages(scope: QueueTicket[], now: Date): { label: string; count: number; ids: number[] }[] {
  const out = AGE_BUCKETS.map((b) => ({ label: b.label, count: 0, ids: [] as number[] }));
  scope.filter(isOpen).forEach((t) => {
    const days = (now.getTime() - new Date(t.created).getTime()) / DAY;
    if (isNaN(days)) return;
    const i = AGE_BUCKETS.findIndex((b) => days < b.max);
    out[i].count += 1;
    out[i].ids.push(t.id);
  });
  return out;
}

// ---------------------------------------------------------------- record gaps

export interface Gap {
  label: string;
  have: number;
  of: number;
  note: string;
}

export function gaps(inView: QueueTicket[]): Gap[] {
  const done = inView.filter((t) => t.status === 'Completed');
  return [
    {
      label: 'Tickets with a department',
      have: inView.filter((t) => t.department.trim() !== '').length,
      of: inView.length,
      note: 'The Raise a ticket form fills it from the sign-in.'
    },
    {
      label: 'Tickets with a category',
      have: inView.filter((t) => t.category.trim() !== '').length,
      of: inView.length,
      note: 'The Raise a ticket form always sets one.'
    },
    {
      label: 'Closed tickets with a root cause',
      have: done.filter((t) => t.rootCause.trim() !== '').length,
      of: done.length,
      note: 'Filled in by IT when closing a ticket.'
    }
  ];
}

// ---------------------------------------------------------------- the three sentences on top

export interface Story {
  badge: string;
  kicker: string;
  text: string;
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

export function days1(n: number): string {
  if (n < 1) return `${Math.max(1, Math.round(n * 24))} hours`;
  const r = Math.round(n * 10) / 10;
  return `${r} ${r === 1 ? 'day' : 'days'}`;
}

export function stories(h: Headline, heat: Heat, now: Date): Story[] {
  const out: Story[] = [];
  const diff = h.raisedThisMonth - h.closedThisMonth;
  out.push({
    badge: diff > 0 ? `+${diff}` : diff < 0 ? `−${-diff}` : '=',
    kicker: `${MONTH_NAMES[now.getMonth()]} so far`,
    text:
      `${h.raisedThisMonth} ${h.raisedThisMonth === 1 ? 'ticket' : 'tickets'} in, ${h.closedThisMonth} closed. ` +
      (diff > 0 ? `The pile grew by ${diff}.` : diff < 0 ? `The pile shrank by ${-diff}.` : 'The pile held steady.')
  });
  if (heat.busiest) {
    out.push({
      badge: HEAT_HOURS[heat.busiest.col],
      kicker: 'Busiest moment',
      text: `${slotText(heat.busiest.day, heat.busiest.col)}. Plan cover around it.`
    });
  }
  if (h.medianDays !== undefined) {
    out.push({
      badge: h.medianDays < 1 ? '<1' : String(Math.round(h.medianDays * 10) / 10),
      kicker: 'Speed',
      text: `Half of these tickets were closed within ${days1(h.medianDays)}.`
    });
  }
  return out;
}
