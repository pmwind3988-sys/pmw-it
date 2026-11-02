// How long a ticket took, from raised to completed, in words a person reads:
// "3 days 4 hrs". Only Completed tickets have one; a cancellation is not a fix.

import { QueueTicket } from './ticket';

const HOUR = 60 * 60 * 1000;

/** Hours from raised to completed; its last update stands in when no completed date was recorded. */
export function hoursToComplete(t: QueueTicket): number | undefined {
  if (t.status !== 'Completed') return undefined;
  const start = new Date(t.created).getTime();
  const end = new Date(t.completedDate || t.modified).getTime();
  if (isNaN(start) || isNaN(end)) return undefined;
  return Math.max(0, (end - start) / HOUR);
}

const plural = (n: number, one: string, many: string): string => `${n} ${n === 1 ? one : many}`;

/** "3 days 4 hrs", "1 day", "5 hrs", "Under 1 hr". Rounded to the hour. */
export function durationText(hours: number): string {
  const whole = Math.round(hours);
  if (whole < 1) return 'Under 1 hr';
  const days = Math.floor(whole / 24);
  const hrs = whole % 24;
  if (days === 0) return plural(hrs, 'hr', 'hrs');
  if (hrs === 0) return plural(days, 'day', 'days');
  return `${plural(days, 'day', 'days')} ${plural(hrs, 'hr', 'hrs')}`;
}

/** The table's and the export's cell: blank for anything not completed. */
export function timeToCompleteText(t: QueueTicket): string {
  const h = hoursToComplete(t);
  return h === undefined ? '' : durationText(h);
}
