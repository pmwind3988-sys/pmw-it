// The date choices on the ticket queue, turned into a from/to pair.
//
// Everything is the browser's LOCAL calendar (Malaysia, for everyone using
// this), never UTC: "today" at 7am in Ipoh is still today, and a UTC day would
// start at 8am and put this morning's tickets under yesterday.
// `from` is inclusive, `to` is exclusive, both instants.

export type DatePreset =
  | 'all'
  | 'today'
  | 'thisWeek'
  | 'pastWeek'
  | 'last30'
  | 'thisMonth'
  | 'thisYear'
  | 'month'
  | 'range';

export const DATE_PRESETS: { id: DatePreset; label: string }[] = [
  { id: 'all', label: 'Any time' },
  { id: 'today', label: 'Today' },
  { id: 'thisWeek', label: 'This week' },
  { id: 'pastWeek', label: 'Past 7 days' },
  { id: 'last30', label: 'Past 30 days' },
  { id: 'thisMonth', label: 'This month' },
  { id: 'thisYear', label: 'This year' },
  { id: 'month', label: 'A specific month' },
  { id: 'range', label: 'Between two dates' }
];

export interface DateChoice {
  preset: DatePreset;
  /** 'YYYY-MM', for preset 'month'. */
  month?: string;
  /** 'YYYY-MM-DD', both inclusive, for preset 'range'. */
  start?: string;
  end?: string;
}

export interface Range {
  from?: Date;
  to?: Date;
}

const startOfDay = (d: Date): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addDays = (d: Date, n: number): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);

/** 'YYYY-MM-DD' as LOCAL midnight. `new Date('2026-09-01')` is UTC midnight, the previous day's 8am here. */
export function parseDay(s: string | undefined): Date | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s ?? '');
  if (!m) return undefined;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return d.getMonth() === Number(m[2]) - 1 ? d : undefined;
}

export function rangeOf(choice: DateChoice, now: Date = new Date()): Range {
  const today = startOfDay(now);
  switch (choice.preset) {
    case 'today':
      return { from: today, to: addDays(today, 1) };
    case 'thisWeek': {
      // Weeks start on Monday.
      const back = (today.getDay() + 6) % 7;
      const monday = addDays(today, -back);
      return { from: monday, to: addDays(monday, 7) };
    }
    case 'pastWeek':
      return { from: addDays(today, -6), to: addDays(today, 1) };
    case 'last30':
      return { from: addDays(today, -29), to: addDays(today, 1) };
    case 'thisMonth':
      return { from: new Date(now.getFullYear(), now.getMonth(), 1), to: new Date(now.getFullYear(), now.getMonth() + 1, 1) };
    case 'thisYear':
      return { from: new Date(now.getFullYear(), 0, 1), to: new Date(now.getFullYear() + 1, 0, 1) };
    case 'month': {
      const m = /^(\d{4})-(\d{2})$/.exec(choice.month ?? '');
      if (!m) return {};
      const y = Number(m[1]);
      const mo = Number(m[2]) - 1;
      return { from: new Date(y, mo, 1), to: new Date(y, mo + 1, 1) };
    }
    case 'range': {
      let start = parseDay(choice.start);
      let end = parseDay(choice.end);
      // Picked back to front is still the same two days.
      if (start && end && start > end) [start, end] = [end, start];
      return { from: start, to: end ? addDays(end, 1) : undefined };
    }
    default:
      return {};
  }
}

export function inRange(when: Date | string, r: Range): boolean {
  const t = typeof when === 'string' ? new Date(when) : when;
  if (isNaN(t.getTime())) return !r.from && !r.to;
  if (r.from && t < r.from) return false;
  if (r.to && t >= r.to) return false;
  return true;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const day = (d: Date): string => `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;

/** What the choice covers, in words, for the screen and the export's file name. */
export function describe(choice: DateChoice, now: Date = new Date()): string {
  const r = rangeOf(choice, now);
  if (choice.preset === 'all') return 'Any time';
  if (!r.from && !r.to) return 'Pick a date';
  if (choice.preset === 'month' && r.from) return `${MONTHS[r.from.getMonth()]} ${r.from.getFullYear()}`;
  if (choice.preset === 'thisYear' && r.from) return String(r.from.getFullYear());
  const last = r.to ? addDays(r.to, -1) : undefined;
  if (r.from && last && r.from.getTime() === last.getTime()) return day(r.from);
  if (r.from && last) return `${day(r.from)} – ${day(last)}`;
  if (r.from) return `From ${day(r.from)}`;
  return `Up to ${day(last as Date)}`;
}
