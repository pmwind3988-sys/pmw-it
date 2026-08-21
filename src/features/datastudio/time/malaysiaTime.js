// Malaysia (Asia/Kuala_Lumpur) time helpers for Data Studio.
//
// Malaysia has used a flat UTC+8 offset with no daylight saving since 1982,
// so this module can use a constant offset instead of timezone-database
// lookups for the arithmetic paths. `formatMYT` still delegates to the ICU
// `Asia/Kuala_Lumpur` timezone via `Intl.DateTimeFormat` for display, which
// is equivalent but keeps locale-correct part ordering/padding for free.

export const MYT_OFFSET_MIN = 480;

// Excel stores dates as a serial day count from 1899-12-31, with the
// fractional part representing time of day. Lotus 1-2-3 (which Excel's
// serial system originates from) incorrectly treats 1900 as a leap year,
// so Excel inherited a phantom 1900-02-29 (serial 60). Serials 1-59 are
// correct as-is (1 = 1900-01-01); serial 60 is the nonexistent leap day;
// serials 61+ need no further correction once you anchor them from
// 1899-12-31, because the standard `(serial - 25569) * 86400000` epoch-day
// formula already assumes a "correct" calendar. To reconcile: for
// serial < 61, add back one day (86400000ms) to undo the phantom leap day
// that the standard formula implicitly subtracts.
const MS_PER_DAY = 86400000;
const EXCEL_EPOCH_OFFSET_DAYS = 25569; // days between 1899-12-31 and 1970-01-01

export function excelSerialToEpochMs(serial) {
  let ms = (serial - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY;
  if (serial < 61) {
    ms += MS_PER_DAY;
  }
  return ms;
}

const ISO_SHAPE = /^\d{4}-\d{2}-\d{2}/;

function splitDateComponents(value) {
  const match = /^(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})/.exec(value);
  if (!match) return null;
  return [match[1], match[2], match[3]];
}

// Scans every value, tracking whether dmy or mdy (or both) is proven by at
// least one value whose first or second component can only be a day
// (i.e. > 12). ISO-shaped values short-circuit as unambiguous.
export function detectDateOrder(values) {
  let provenDmy = false;
  let provenMdy = false;

  for (const raw of values) {
    const value = String(raw ?? '').trim();
    if (!value) continue;

    if (ISO_SHAPE.test(value)) {
      return 'iso';
    }

    const parts = splitDateComponents(value);
    if (!parts) continue;

    const first = Number(parts[0]);
    const second = Number(parts[1]);

    if (first > 12) provenDmy = true;
    if (second > 12) provenMdy = true;
  }

  if (provenDmy && provenMdy) return 'conflict';
  if (provenDmy) return 'dmy';
  if (provenMdy) return 'mdy';
  return 'ambiguous';
}

// Parses "D/M/Y" or "M/D/Y" (also accepting "-" separators) with an
// optional trailing "HH:mm" (or "HH:mm:ss") time-of-day component.
const DATE_TIME_RE =
  /^(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;

function parseOrderedDateString(value, order) {
  const match = DATE_TIME_RE.exec(value.trim());
  if (!match) return NaN;

  const [, a, b, c, hh, mm, ss] = match;
  let day;
  let month;
  let year;

  if (order === 'iso') {
    year = Number(a);
    month = Number(b);
    day = Number(c);
  } else if (order === 'mdy') {
    month = Number(a);
    day = Number(b);
    year = Number(c);
  } else {
    // dmy (default)
    day = Number(a);
    month = Number(b);
    year = Number(c);
  }

  if (year < 100) year += 2000;

  const hours = hh ? Number(hh) : 0;
  const minutes = mm ? Number(mm) : 0;
  const seconds = ss ? Number(ss) : 0;

  const ms = Date.UTC(year, month - 1, day, hours, minutes, seconds);
  if (Number.isNaN(ms)) return NaN;

  // Round-trip validation: Date.UTC silently rolls invalid components over
  // into the next day/month/hour instead of rejecting them (e.g. day=32
  // becomes the 1st of the following month). Reading the components back
  // off the constructed instant and comparing them to what was parsed
  // catches any date that does not actually exist -- e.g. 31 February, or
  // 29 February in a non-leap year -- and does so before the caller (in
  // toEpochMs) applies any sourceZone shift, so a valid late-evening time
  // that crosses midnight after shifting is never mistaken for one of
  // these.
  const check = new Date(ms);
  if (
    check.getUTCFullYear() !== year ||
    check.getUTCMonth() !== month - 1 ||
    check.getUTCDate() !== day ||
    check.getUTCHours() !== hours ||
    check.getUTCMinutes() !== minutes ||
    check.getUTCSeconds() !== seconds
  ) {
    return NaN;
  }

  return ms;
}

export function toEpochMs(value, opts = {}) {
  const { order = 'dmy', sourceZone = 'local', dateOnly = false } = opts;

  if (value instanceof Date) {
    const t = value.getTime();
    return Number.isNaN(t) ? NaN : t;
  }

  if (typeof value !== 'string' || value.trim() === '') {
    return NaN;
  }

  let ms = parseOrderedDateString(value, order);
  if (Number.isNaN(ms)) return NaN;

  if (!dateOnly && sourceZone === 'utc') {
    ms += MYT_OFFSET_MIN * 60000;
  }

  return ms;
}

// Builds a `byType` lookup (day/month/year/hour/minute -> string) from
// Intl.DateTimeFormat#formatToParts, rather than calling `.format()` and
// hoping. This matters because `.format()` with both date and time parts
// renders `en-GB` as "15/01/2024, 08:00" -- with a comma -- and we need
// the separators under our own control.
function getPartsMYT(epochMs, options) {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kuala_Lumpur',
    // Explicit on purpose, and deliberately NOT paired with `hour12: false`:
    // per the Intl.DateTimeFormat spec, an explicit `hour12` (true or
    // false) overrides/nullifies any `hourCycle` option entirely -- verified
    // on this engine, {hour12:false, hourCycle:'h24'} still resolves to
    // 'h23', proving hourCycle is silently discarded whenever hour12 is
    // also present. So on an engine/ICU version whose bare `hour12:false`
    // locale default is 'h24' (rendering midnight as "24:00" -- the exact
    // browser-observed bug this guards against), adding hourCycle:'h23'
    // alongside hour12:false would NOT fix it; only hourCycle on its own
    // reliably pins the 00-23 range. Do not "simplify" this by adding back
    // `hour12: false` or removing `hourCycle` -- either one reopens the bug.
    hourCycle: 'h23',
    ...options,
  });
  const byType = {};
  for (const part of formatter.formatToParts(new Date(epochMs))) {
    byType[part.type] = part.value;
  }
  return byType;
}

/**
 * The twelve-hour twin of getPartsMYT, and it follows the same discipline for
 * the same reason: pin the hour cycle, never pass `hour12`. 'h12' is the 1-12
 * cycle that renders midnight as "12 AM"; 'h11' is the 0-11 cycle that would
 * render it as "0 AM" -- the exact mirror of the "24:00" bug its sibling
 * guards against. Do not "simplify" this by adding `hour12: true`.
 */
function getParts12MYT(epochMs, options) {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kuala_Lumpur',
    hourCycle: 'h12',
    ...options,
  });
  const byType = {};
  for (const part of formatter.formatToParts(new Date(epochMs))) {
    byType[part.type] = part.value;
  }
  return byType;
}

export function formatMYT(epochMs, style = 'datetime') {
  if (Number.isNaN(epochMs)) return '—';

  const datePart = () => {
    const { day, month, year } = getPartsMYT(epochMs, {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
    return `${day}/${month}/${year}`;
  };

  const timePart = () => {
    const { hour, minute } = getPartsMYT(epochMs, {
      hour: '2-digit',
      minute: '2-digit',
    });
    return `${hour}:${minute}`;
  };

  const timePart12 = () => {
    const { hour, minute, dayPeriod } = getParts12MYT(epochMs, {
      hour: '2-digit',
      minute: '2-digit',
    });
    // Some ICU builds separate the day period with a narrow no-break space.
    return `${hour}:${minute} ${dayPeriod.replace(/\s/g, '').toUpperCase()}`;
  };

  if (style === 'date') return datePart();
  if (style === 'time') return timePart();
  if (style === 'time12') return timePart12();
  if (style === 'datetime12') return `${datePart()} ${timePart12()}`;
  return `${datePart()} ${timePart()}`;
}
