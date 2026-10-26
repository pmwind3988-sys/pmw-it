import { MYT_OFFSET_MIN, toEpochMs } from '../../utils/malaysiaTime.js';

/**
 * Dates as the machine page words them: `21 Jan 2027`, never `2027-01-21` and
 * never a US-style month first. Pure: no Intl, so the output does not depend on
 * the browser's locale or timezone.
 */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** `2027-01-21` (a stored date-only value) to `21 Jan 2027`; null when it is not one. */
export function dayLabel(iso) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso ?? '').trim());
  if (!match) return null;
  const month = Number(match[2]);
  if (month < 1 || month > 12) return null;
  return `${Number(match[3])} ${MONTHS[month - 1]} ${match[1]}`;
}

/** An instant (epoch ms, or a value toEpochMs reads) as the Malaysia calendar day. */
export function scanDayLabel(value) {
  let ms;
  try {
    ms = typeof value === 'number' ? value : toEpochMs(value);
  } catch {
    return null;
  }
  if (!Number.isFinite(ms)) return null;
  const local = new Date(ms + MYT_OFFSET_MIN * 60000);
  return `${local.getUTCDate()} ${MONTHS[local.getUTCMonth()]} ${local.getUTCFullYear()}`;
}
