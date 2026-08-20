import { describe, it, expect } from 'vitest';
import {
  MYT_OFFSET_MIN,
  excelSerialToEpochMs,
  detectDateOrder,
  toEpochMs,
  formatMYT,
} from './malaysiaTime.js';

describe('MYT_OFFSET_MIN', () => {
  it('is a flat UTC+8 with no DST', () => {
    expect(MYT_OFFSET_MIN).toBe(480);
  });
});

describe('excelSerialToEpochMs', () => {
  // Excel's epoch is 1899-12-31, and it wrongly believes 1900 was a leap year.
  // Serial 60 is 1900-02-29 -- a date that never existed.
  it('converts serial 61 (1900-03-01) correctly', () => {
    expect(excelSerialToEpochMs(61)).toBe(Date.UTC(1900, 2, 1));
  });

  it('converts serials below 61 using the pre-bug offset', () => {
    expect(excelSerialToEpochMs(1)).toBe(Date.UTC(1900, 0, 1));
  });

  it('converts a modern serial', () => {
    expect(excelSerialToEpochMs(45000)).toBe(Date.UTC(2023, 2, 15));
  });

  it('carries the fractional part as time of day', () => {
    expect(excelSerialToEpochMs(45000.5)).toBe(Date.UTC(2023, 2, 15, 12));
  });
});

describe('detectDateOrder', () => {
  it('proves dmy when a first component exceeds 12', () => {
    expect(detectDateOrder(['13/01/2024', '05/02/2024'])).toBe('dmy');
  });

  it('proves mdy when a second component exceeds 12', () => {
    expect(detectDateOrder(['01/13/2024', '02/05/2024'])).toBe('mdy');
  });

  it('reports a conflict when both are proven', () => {
    expect(detectDateOrder(['13/01/2024', '01/13/2024'])).toBe('conflict');
  });

  it('reports ambiguous when nothing proves it', () => {
    expect(detectDateOrder(['01/02/2024', '03/04/2024'])).toBe('ambiguous');
  });

  it('recognises ISO as unambiguous', () => {
    expect(detectDateOrder(['2024-01-13', '2024-02-05'])).toBe('iso');
  });
});

describe('toEpochMs', () => {
  it('passes Date objects straight through', () => {
    const d = new Date(Date.UTC(2024, 0, 15, 3, 30));
    expect(toEpochMs(d)).toBe(d.getTime());
  });

  it('reads dmy strings when told to', () => {
    expect(toEpochMs('15/01/2024', { order: 'dmy' }))
      .toBe(Date.UTC(2024, 0, 15));
  });

  it('reads the same string differently as mdy', () => {
    expect(toEpochMs('05/01/2024', { order: 'mdy' }))
      .toBe(Date.UTC(2024, 4, 1));
  });

  it('does not shift when the source is local', () => {
    expect(toEpochMs('15/01/2024 08:00', { order: 'dmy', sourceZone: 'local' }))
      .toBe(Date.UTC(2024, 0, 15, 8, 0));
  });

  it('shifts by +8 when the source is UTC', () => {
    expect(toEpochMs('15/01/2024 08:00', { order: 'dmy', sourceZone: 'utc' }))
      .toBe(Date.UTC(2024, 0, 15, 16, 0));
  });

  // Spec §9 hard rule -- shifting a pure date moves it to the wrong day.
  it('never shifts a date-only column even when marked UTC', () => {
    expect(toEpochMs('15/01/2024', { order: 'dmy', sourceZone: 'utc', dateOnly: true }))
      .toBe(Date.UTC(2024, 0, 15));
  });

  it('returns NaN for unparseable input', () => {
    expect(toEpochMs('not a date')).toBeNaN();
    expect(toEpochMs(null)).toBeNaN();
    expect(toEpochMs('')).toBeNaN();
  });

  it('rejects a day that does not exist in that month', () => {
    expect(toEpochMs('31/02/2024', { order: 'dmy' })).toBeNaN();
    expect(toEpochMs('32/01/2024', { order: 'dmy' })).toBeNaN();
  });

  it('accepts a real leap day', () => {
    expect(toEpochMs('29/02/2024', { order: 'dmy' })).toBe(Date.UTC(2024, 1, 29));
  });

  it('rejects a leap day in a non-leap year', () => {
    expect(toEpochMs('29/02/2023', { order: 'dmy' })).toBeNaN();
  });

  it('still shifts a valid late-evening time across midnight', () => {
    expect(toEpochMs('15/01/2024 23:30', { order: 'dmy', sourceZone: 'utc' }))
      .toBe(Date.UTC(2024, 0, 16, 7, 30));
  });

  it('rejects an out-of-range month', () => {
    expect(toEpochMs('13/01/2024', { order: 'mdy' })).toBeNaN();
  });

  it('parses a plain ISO date (YYYY-MM-DD)', () => {
    expect(toEpochMs('2024-01-15', { order: 'iso' })).toBe(Date.UTC(2024, 0, 15));
  });

  it('parses a full ISO datetime with a T separator', () => {
    expect(toEpochMs('2024-01-15T08:00', { order: 'iso' }))
      .toBe(Date.UTC(2024, 0, 15, 8, 0));
  });

  it('round-trip validation rejects an ISO date that does not exist', () => {
    expect(toEpochMs('2024-02-31', { order: 'iso' })).toBeNaN();
  });

  it('never shifts a date-only ISO value even when marked UTC', () => {
    expect(toEpochMs('2024-01-15', { order: 'iso', sourceZone: 'utc', dateOnly: true }))
      .toBe(Date.UTC(2024, 0, 15));
  });

  it('uplifts a 2-digit year into the 2000s', () => {
    expect(toEpochMs('15/01/24', { order: 'dmy' })).toBe(Date.UTC(2024, 0, 15));
  });

  // Decision: the uplift is unconditional (+2000 for any year < 100) -- no
  // Excel-style century windowing (e.g. 00-29 -> 2000s, 30-99 -> 1900s).
  // '99' therefore resolves to 2099, not 1999. See task-1-report.md for why
  // this was kept as-is rather than windowed.
  it('uplifts a 2-digit year of 99 to 2099 (no century windowing)', () => {
    expect(toEpochMs('15/01/99', { order: 'dmy' })).toBe(Date.UTC(2099, 0, 15));
  });
});

describe('formatMYT', () => {
  it('formats as DD/MM/YYYY HH:mm in Malaysian local time', () => {
    // 2024-01-15T00:00Z is 08:00 on the 15th in KL.
    expect(formatMYT(Date.UTC(2024, 0, 15, 0, 0))).toBe('15/01/2024 08:00');
  });

  it('formats date-only style without a time', () => {
    expect(formatMYT(Date.UTC(2024, 0, 15, 0, 0), 'date')).toBe('15/01/2024');
  });

  it('formats time-only style without a date', () => {
    expect(formatMYT(Date.UTC(2024, 0, 15, 0, 0), 'time')).toBe('08:00');
  });

  it('formats midnight in MYT as 00:00, not 24:00', () => {
    // 2024-01-14T16:00Z is 00:00 on the 15th in KL -- the ICU hour=24 edge case.
    expect(formatMYT(Date.UTC(2024, 0, 14, 16, 0), 'time')).toBe('00:00');
  });

  it('renders NaN as an em dash rather than "Invalid Date"', () => {
    expect(formatMYT(NaN)).toBe('—');
  });
});
