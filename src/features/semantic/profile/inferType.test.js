import { describe, it, expect } from 'vitest';
import {
  isNullish, parseNumberLike, parseBooleanLike, inferType, MIN_IDENTIFIER_RUN,
} from './inferType.js';

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

  // The leading-zero guard has to survive currency, parenthesis and sign
  // stripping -- otherwise a cost centre written '$007' is silently 7.
  it.each(['$007', 'RM 007', '(007)', '-007', '+007'])(
    'refuses %j, whose leading zero only shows after stripping', (input) => {
      expect(parseNumberLike(input).ok).toBe(false);
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

  // I3 -- clause (b) of the identifier override (distinct ratio > 0.95 AND
  // an identifier-ish name) had zero coverage: every existing identifier
  // fixture was a consecutive run, so clause (a) fired first. These two
  // pin the boundary in both directions with NON-consecutive values.
  it('detects a non-consecutive unique column as an identifier when the name says so', () => {
    const values = ['101', '207', '313', '429', '555', '661'];
    expect(inferType(values, 'Ref No')).toMatchObject({
      type: 'identifier', role: 'ignored',
    });
  });

  it('leaves the same non-consecutive values as a measure under a measure name', () => {
    const values = ['101', '207', '313', '429', '555', '661'];
    expect(inferType(values, 'Amount')).toMatchObject({
      type: 'numeric', role: 'measure',
    });
  });

  // The brief's literal /id|no|code|ref|serial/i is an unanchored substring
  // match, so 'Paid Amount' (contains 'id') was demoted to role 'ignored'.
  // Matching whole tokens keeps the intended hits and drops the collisions.
  it.each(['Paid Amount', 'Width', 'Notes', 'Income', 'Humidity'])(
    'does not mistake %j for an identifier name', (name) => {
      const values = ['101', '207', '313', '429', '555', '661'];
      expect(inferType(values, name).type).toBe('numeric');
    });

  it.each(['Employee ID', 'EmpID', 'Emp_ID', 'Serial Number', 'Cost Code', 'Ref'])(
    'still reads %j as an identifier name', (name) => {
      const values = ['101', '207', '313', '429', '555', '661'];
      expect(inferType(values, name).type).toBe('identifier');
    });

  // A two-value run is not evidence of a row number -- it is two ordinary
  // measurements just as readily. Short runs must not fire clause (a).
  it('does not call a two-value ascending run an identifier', () => {
    expect(inferType(['10', '11'], 'Delta').type).toBe('numeric');
  });

  it('still calls a long consecutive run an identifier', () => {
    const values = Array.from({ length: MIN_IDENTIFIER_RUN }, (_, i) => String(i + 1));
    expect(inferType(values, 'Delta').type).toBe('identifier');
  });

  // I4.1 -- ZERO_WIDTH_RE is load-bearing in a way NBSP handling is not:
  // U+200B-U+200D and U+FEFF are NOT in JavaScript's \s class, so nothing
  // else strips them. Escape sequences, never literals -- a literal
  // invisible character does not survive retyping and voids the test.
  it('strips zero-width characters before comparing values', () => {
    const v = inferType(['HR', 'H\u200BR', 'HR', '\uFEFFHR'], 'Department');
    expect(v.distinctCount).toBe(1);
  });

  it('parses a number that is peppered with zero-width characters', () => {
    expect(parseNumberLike('1\u200B234').value).toBe(1234);
  });

  // I4.2 -- casualties must hold the RAW value, not a normalised one, so
  // the user sees what is actually in their sheet. The pre-existing
  // fixture used 'pending', identical raw and normalised, so a mutation
  // to normalise them failed no test.
  it('reports casualties with their raw surrounding whitespace intact', () => {
    const v = inferType([...Array(19).fill('10'), '  pending  '], 'Amount');
    expect(v.casualties).toContain('  pending  ');
  });

  // I4.3 -- a column is isPercent only when EVERY numeric match was a
  // percent; `percentMatches > 0` failed no test before this.
  it('flags a column as percent only when every value is a percent', () => {
    expect(inferType(['45%', '50%', '55%'], 'Rate').isPercent).toBe(true);
    expect(inferType(['45%', '50', '55%'], 'Rate').isPercent).toBe(false);
  });

  // I4.4 -- `type: 'date'` is a whole member of the verdict union and the
  // flag `dateOnly` consumers key off, but nothing asserted it; hardcoding
  // 'datetime' failed no test.
  it('distinguishes a date-only column from one carrying a time', () => {
    expect(inferType(['13/01/2024', '05/02/2024'], 'Join Date').type).toBe('date');
    expect(inferType(['13/01/2024 08:30', '05/02/2024 09:00'], 'Join Date').type)
      .toBe('datetime');
  });

  // I4.5 -- distinctCount is computed over NORMALISED values, so padding
  // does not invent extra categories.
  it('counts distinct values after normalising them', () => {
    const v = inferType(['HR', ' HR ', 'IT', 'IT '], 'Department');
    expect(v.distinctCount).toBe(2);
  });

  // A pure-Date column has no strings to be ambiguous about, so `null` is
  // the accurate dateOrder rather than a guess.
  it('reports no date order for a column of Date objects', () => {
    const values = [new Date(Date.UTC(2024, 0, 1)), new Date(Date.UTC(2024, 0, 2))];
    expect(inferType(values, 'Created').dateOrder).toBe(null);
  });
});

describe('multi-select detection', () => {
  const survey = [
    'Data Collection;Data Cleaning;Report Generation;',
    'Data Collection;Approval Tracking;',
    'Report Generation;Data Collection;',
    'Approval Tracking;',
    'Data Cleaning;Report Generation;',
    'Data Collection;Data Cleaning;',
  ];

  it('reads a semicolon-joined multi-select as multi', () => {
    const verdict = inferType(survey, 'Which challenges');
    expect(verdict.type).toBe('multi');
    expect(verdict.role).toBe('dimension');
    expect(verdict.separator).toBe(';');
  });

  it('leaves an ordinary categorical column alone', () => {
    const verdict = inferType(['IT', 'Finance', 'IT', 'Logistics', 'IT'], 'Department');
    expect(verdict.type).toBe('categorical');
  });

  it('does not read prose containing semicolons as multi', () => {
    // Parts that never repeat are sentences, not options.
    const prose = [
      'The process is manual; it takes hours to reconcile every month.',
      'We chase approvals by email; nobody knows the current status.',
      'Reports are rebuilt from scratch; version control is guesswork.',
      'Files live in five places; finding the latest one is luck.',
      'Data is retyped between systems; typos are common.',
      'Updates arrive on WhatsApp; important ones get missed.',
    ];
    expect(inferType(prose, 'Describe').type).not.toBe('multi');
  });

  it('does not read a single-option column as multi', () => {
    const single = ['IT;', 'Finance;', 'IT;', 'Logistics;', 'IT;', 'Finance;'];
    expect(inferType(single, 'Department').type).not.toBe('multi');
  });
});
