import { describe, it, expect } from 'vitest';
import { isNullish, parseNumberLike, parseBooleanLike, inferType } from './inferType.js';

describe('isNullish', () => {
  it.each([['', true], ['-', true], ['N/A', true], ['#DIV/0!', true],
          ['NIL', true], ['  ', true], ['0', false], ['abc', false]])(
    'treats %j as nullish=%s', (input, expected) => {
      expect(isNullish(input)).toBe(expected);
    });
});

describe('parseNumberLike', () => {
  it('strips thousands separators', () => {
    expect(parseNumberLike('1,234.50')).toMatchObject({ ok: true, value: 1234.5 });
  });

  it('strips a currency prefix', () => {
    expect(parseNumberLike('RM 1,234')).toMatchObject({ ok: true, value: 1234 });
    expect(parseNumberLike('$1,234')).toMatchObject({ ok: true, value: 1234 });
  });

  it('reads percentages as fractions and flags them', () => {
    expect(parseNumberLike('45%')).toMatchObject({
      ok: true, value: 0.45, isPercent: true,
    });
  });

  it('reads accounting negatives', () => {
    expect(parseNumberLike('(1,234)')).toMatchObject({ ok: true, value: -1234 });
  });

  it('strips non-breaking spaces', () => {
    // Written as an escape sequence on purpose: a literal U+00A0 is
    // invisible in a diff and does not survive being retyped, which
    // would silently turn this test into a no-op. Covers both the NBSP
    // case and the ordinary-space case (a European thousands separator).
    expect(parseNumberLike('1\u00A0234')).toMatchObject({ ok: true, value: 1234 });
    expect(parseNumberLike('1 234')).toMatchObject({ ok: true, value: 1234 });
  });

  // Spec §7.4 -- this rule protects employee IDs and cost centres.
  it('refuses values with leading zeros', () => {
    expect(parseNumberLike('007').ok).toBe(false);
  });

  it('accepts a bare zero', () => {
    expect(parseNumberLike('0')).toMatchObject({ ok: true, value: 0 });
  });

  it('rejects non-numeric text', () => {
    expect(parseNumberLike('pending').ok).toBe(false);
  });

  it('cancels a double negative from accounting parentheses', () => {
    expect(parseNumberLike('(-5)')).toMatchObject({ ok: true, value: 5 });
  });
});

describe('parseBooleanLike', () => {
  it.each([['yes', true], ['NO', false], ['true', true], ['Y', true], ['n', false]])(
    'reads %j', (input, expected) => {
      expect(parseBooleanLike(input)).toMatchObject({ ok: true, value: expected });
    });

  // Spec §7.3 -- 0/1 is numeric far more often than it is boolean.
  it('refuses 0 and 1', () => {
    expect(parseBooleanLike('1').ok).toBe(false);
    expect(parseBooleanLike('0').ok).toBe(false);
  });
});

describe('inferType', () => {
  it('types a clean numeric column as a measure', () => {
    const v = inferType(['10', '20', '30', '40'], 'Amount');
    expect(v).toMatchObject({ type: 'numeric', role: 'measure' });
  });

  // Spec §7.2 -- the 95% rule, and its casualties must be reported.
  it('still types numeric at 95% and reports the casualties', () => {
    const values = [...Array(19).fill('10'), 'pending'];
    const v = inferType(values, 'Amount');
    expect(v.type).toBe('numeric');
    expect(v.casualtyCount).toBe(1);
    expect(v.casualties).toContain('pending');
  });

  it('falls back to categorical below 95%', () => {
    const values = [...Array(17).fill('10'), 'a', 'b', 'c'];
    expect(inferType(values, 'Amount').type).toBe('categorical');
  });

  it('ignores nulls when computing the ratio', () => {
    const v = inferType(['10', '20', 'N/A', '', '-'], 'Amount');
    expect(v).toMatchObject({ type: 'numeric', nullCount: 3 });
  });

  it('types a low-cardinality string column as a dimension', () => {
    const v = inferType(['HR', 'IT', 'HR', 'Finance', 'IT'], 'Department');
    expect(v).toMatchObject({ type: 'categorical', role: 'dimension', distinctCount: 3 });
  });

  it('types high-cardinality free text as ignored', () => {
    const values = Array.from({ length: 200 }, (_, i) => `remark number ${i}`);
    expect(inferType(values, 'Remarks')).toMatchObject({ type: 'text', role: 'ignored' });
  });

  it('types a known boolean pair as a dimension', () => {
    const v = inferType(['Yes', 'No', 'Yes', 'No'], 'Active');
    expect(v).toMatchObject({ type: 'boolean', role: 'dimension' });
  });

  // Spec §7.3 -- summing employee IDs is meaningless.
  it('detects an identifier by name and uniqueness, and refuses to call it a measure', () => {
    const values = Array.from({ length: 50 }, (_, i) => String(1000 + i));
    const v = inferType(values, 'Employee ID');
    expect(v).toMatchObject({ type: 'identifier', role: 'ignored' });
  });

  it('detects a monotonic integer sequence as an identifier even without a matching name', () => {
    const values = Array.from({ length: 50 }, (_, i) => String(i + 1));
    expect(inferType(values, 'Seq').type).toBe('identifier');
  });

  it('types Date objects as datetime', () => {
    const values = [new Date(Date.UTC(2024, 0, 1)), new Date(Date.UTC(2024, 0, 2))];
    expect(inferType(values, 'Created')).toMatchObject({
      type: 'datetime', role: 'temporal',
    });
  });

  it('carries the detected date order onto the verdict', () => {
    const v = inferType(['13/01/2024', '05/02/2024'], 'Join Date');
    expect(v).toMatchObject({ role: 'temporal', dateOrder: 'dmy' });
  });

  it('leaves a conflicting date column as text for the user to resolve', () => {
    const v = inferType(['13/01/2024', '01/13/2024'], 'Join Date');
    expect(v).toMatchObject({ type: 'text', dateOrder: 'conflict' });
  });

  // A date-order conflict must only block the column when the datetime
  // candidate would otherwise have won. Here two stray conflicting dates
  // sit inside an otherwise-numeric column that clears the 95% bar, so
  // the conflict is informational only -- the column should still type
  // as numeric, not get discarded as text.
  it('lets a numeric column win despite two conflicting stray date values, but reports the conflict', () => {
    const values = [...Array(100).fill('10'), '13/01/2024', '01/13/2024'];
    const v = inferType(values, 'Amount');
    expect(v).toMatchObject({ type: 'numeric', role: 'measure', dateOrder: 'conflict' });
  });

  // C1 -- ISO-8601 is the most common machine-export date format, and the
  // 'iso' dateOrder was previously mapped to 'dmy' before reaching
  // toEpochMs, making every ISO date column misclassify as categorical.
  it('types an ISO date column as temporal with dateOrder iso', () => {
    const v = inferType(['2024-01-15', '2024-02-20', '2024-03-25'], 'Join Date');
    expect(v).toMatchObject({ type: 'date', role: 'temporal', dateOrder: 'iso' });
  });

  it('types an ISO datetime column as temporal datetime with dateOrder iso', () => {
    const v = inferType(['2024-01-15T08:30:00', '2024-02-20T09:00:00'], 'Created');
    expect(v).toMatchObject({ type: 'datetime', role: 'temporal', dateOrder: 'iso' });
  });

  it('types an all-null column as empty', () => {
    expect(inferType(['', 'N/A', '-'], 'Blank')).toMatchObject({
      type: 'empty', role: 'ignored',
    });
  });

  it('preserves leading-zero codes as categorical, not numeric', () => {
    const v = inferType(['007', '008', '007', '009'], 'Cost Centre');
    expect(v.type).toBe('categorical');
  });
});
