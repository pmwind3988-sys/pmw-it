// Column type inference for Data Studio -- spec §7.2-7.4.
//
// Given a column's raw values (as they come off a parsed spreadsheet) and
// its header name, decides what kind of thing the column is (numeric,
// date, boolean, categorical, free text, identifier, or empty) and what
// role it should play in charts (measure, dimension, temporal, ignored).
//
// The guiding principle throughout is: silent coercion is the bug this
// module exists to prevent. When a value doesn't fit the winning type, it
// is reported as a "casualty" rather than quietly mangled into one.

import { toEpochMs, detectDateOrder } from '../time/malaysiaTime.js';

export const NULL_TOKENS = new Set([
  '', '-', '–', '—', 'n/a', 'na', 'null', 'nil', 'tbd',
  '#n/a', '#div/0!', '#ref!', '#value!', '#name?',
]);

// Zero-width characters that are invisible in a diff but poison parsing:
// ZWSP (U+200B), ZWNJ (U+200C), ZWJ (U+200D), BOM/ZWNBSP (U+FEFF).
// Non-breaking space (U+00A0) is handled separately since it collapses to
// a regular space rather than vanishing. Written as escape sequences
// rather than literal characters on purpose: a literal invisible
// character in source does not survive being retyped or diffed and can
// silently rot into a no-op.
const ZERO_WIDTH_RE = /[\u200B-\u200D\uFEFF]/g;
const NBSP_RE = /\u00A0/g;

// Normalises a raw cell value to a trimmed string for comparison/parsing.
// `Date` objects are the one exception -- callers must check for those
// before calling this, since stringifying a Date here would destroy it.
function normalizeToString(value) {
  const s = String(value ?? '');
  return s.replace(ZERO_WIDTH_RE, '').replace(NBSP_RE, ' ').trim();
}

export function isNullish(value) {
  const s = normalizeToString(value).toLowerCase();
  return NULL_TOKENS.has(s);
}

const LEADING_ZERO_RE = /^0\d/;
const CURRENCY_PREFIX_RE = /^(RM|USD|MYR|\$|€|£|¥)\s*/i;
const ACCOUNTING_NEGATIVE_RE = /^\((.*)\)$/;

// Parses a value as a number, tolerant of the messy formats spreadsheets
// throw at us: thousands separators, currency prefixes, percent suffixes,
// accounting-style negatives in parentheses, and non-breaking spaces used
// as thousands separators.
//
// Critically, the leading-zero check runs on the RAW string before any
// stripping happens. '007' must never parse as 7 -- that is what silently
// mangles employee IDs and cost centres into numbers.
export function parseNumberLike(rawValue) {
  const fail = { ok: false, value: NaN, isPercent: false };

  if (rawValue instanceof Date) return fail;

  const raw = String(rawValue ?? '');
  const rawTrimmed = raw.replace(ZERO_WIDTH_RE, '').replace(NBSP_RE, ' ').trim();
  if (rawTrimmed === '') return fail;

  // Leading-zero rejection must run on the string BEFORE stripping
  // currency/percent/separator characters, per spec §7.4.
  if (LEADING_ZERO_RE.test(rawTrimmed)) return fail;

  let s = rawTrimmed;

  // Accounting negatives: "(1,234)" -> negative.
  let negative = false;
  const acctMatch = ACCOUNTING_NEGATIVE_RE.exec(s);
  if (acctMatch) {
    negative = true;
    s = acctMatch[1].trim();
  }

  // Currency prefix.
  s = s.replace(CURRENCY_PREFIX_RE, '');

  // Percent suffix.
  let isPercent = false;
  if (/%$/.test(s)) {
    isPercent = true;
    s = s.slice(0, -1);
  }

  s = s.trim();

  // Leading '-' after stripping currency (e.g. "-RM 5" would be unusual,
  // but "-5" or "- 5" should still work).
  if (/^-/.test(s)) {
    negative = negative || true;
    s = s.slice(1).trim();
  } else if (/^\+/.test(s)) {
    s = s.slice(1).trim();
  }

  // Strip thousands separators (commas) and any remaining whitespace used
  // as a separator.
  s = s.replace(/,/g, '').replace(/\s+/g, '');

  if (s === '' || !/^\d+(\.\d+)?$/.test(s)) return fail;

  let value = Number(s);
  if (!Number.isFinite(value)) return fail;

  if (negative) value = -value;
  if (isPercent) value = value / 100;

  return { ok: true, value, isPercent };
}

const TRUE_WORDS = new Set(['yes', 'true', 'y']);
const FALSE_WORDS = new Set(['no', 'false', 'n']);

// Parses a value as a boolean, but only from known word pairs
// (yes/no, true/false, y/n). Deliberately does NOT accept '0'/'1' --
// those are numeric far more often than boolean (spec §7.3).
export function parseBooleanLike(rawValue) {
  const fail = { ok: false, value: false };
  if (rawValue instanceof Date) return fail;

  const s = normalizeToString(rawValue).toLowerCase();
  if (TRUE_WORDS.has(s)) return { ok: true, value: true };
  if (FALSE_WORDS.has(s)) return { ok: true, value: false };
  return fail;
}

// Checks whether a set of raw string values (already filtered to
// non-null) forms a consecutive integer sequence with a step of exactly
// 1, in the order given. Ties to the binding ruling in the brief: this is
// the ONLY condition that fires the identifier override regardless of
// column name (part (a)); merely-monotonic-but-non-consecutive sequences
// like ['10','20','30','40'] must NOT trigger it.
function isConsecutiveIntegerSequence(rawValues) {
  if (rawValues.length < 2) return false;

  const nums = [];
  for (const raw of rawValues) {
    const s = normalizeToString(raw);
    if (!/^-?\d+$/.test(s)) return false;
    nums.push(Number(s));
  }

  const ascending = nums[1] > nums[0];
  for (let i = 1; i < nums.length; i++) {
    const step = nums[i] - nums[i - 1];
    if (ascending && step !== 1) return false;
    if (!ascending && step !== -1) return false;
  }
  return true;
}

const IDENTIFIER_NAME_RE = /id|no|code|ref|serial/i;

// Whether a value carries an explicit time-of-day component. A native
// `Date` instance always represents a specific millisecond instant -- it
// has no "date-only" form the way a string can lack one -- so any Date
// object counts as carrying a time component, even one that lands
// exactly at UTC midnight. Strings only count when the text itself
// includes an "HH:mm" part; a bare "13/01/2024" does not.
function hasTimeComponent(rawValue) {
  if (rawValue instanceof Date) return true;
  const s = normalizeToString(rawValue);
  return /\d{1,2}:\d{2}/.test(s);
}

function makeEmptyVerdict(nullCount) {
  return {
    type: 'empty',
    role: 'ignored',
    confidence: 0,
    dateOrder: null,
    isPercent: false,
    nullCount,
    distinctCount: 0,
    casualties: [],
    casualtyCount: 0,
  };
}

// Infers the type and chart role of a column from its raw values and
// header name.
//
// Resolution order: boolean -> datetime -> numeric -> categorical -> text,
// then the identifier override runs last over numeric/text/categorical
// verdicts (spec §7.3). First candidate type at >= 95% confidence (over
// non-null values) wins.
export function inferType(values, columnName) {
  const nonNull = [];
  let nullCount = 0;

  for (const v of values) {
    if (v instanceof Date) {
      if (Number.isNaN(v.getTime())) {
        nullCount++;
      } else {
        nonNull.push(v);
      }
      continue;
    }
    if (isNullish(v)) {
      nullCount++;
    } else {
      nonNull.push(v);
    }
  }

  const nonNullCount = nonNull.length;

  if (nonNullCount === 0) {
    return makeEmptyVerdict(nullCount);
  }

  const distinctSet = new Set(nonNull.map((v) => (v instanceof Date ? v.getTime() : normalizeToString(v))));
  const distinctCount = distinctSet.size;

  const THRESHOLD = 0.95;

  // --- boolean -------------------------------------------------------
  {
    const casualties = [];
    let matches = 0;
    for (const v of nonNull) {
      if (parseBooleanLike(v).ok) matches++;
      else if (casualties.length < 5) casualties.push(rawLabel(v));
    }
    const confidence = matches / nonNullCount;
    if (confidence >= THRESHOLD) {
      return {
        type: 'boolean',
        role: 'dimension',
        confidence,
        dateOrder: null,
        isPercent: false,
        nullCount,
        distinctCount,
        casualties,
        casualtyCount: nonNullCount - matches,
      };
    }
  }

  // --- datetime / date -------------------------------------------------
  {
    const stringValues = nonNull.filter((v) => !(v instanceof Date));
    const dateOrder = stringValues.length > 0 ? detectDateOrder(stringValues.map(normalizeToString)) : null;

    // 'conflict' means two values disagree on day/month placement -- the
    // column is left as text for the user to resolve rather than guessing.
    if (dateOrder !== 'conflict') {
      const order = dateOrder === 'iso' || dateOrder === null ? 'dmy' : dateOrder;
      const casualties = [];
      let matches = 0;
      let anyTime = false;

      for (const v of nonNull) {
        if (v instanceof Date) {
          matches++;
          if (hasTimeComponent(v)) anyTime = true;
          continue;
        }
        const s = normalizeToString(v);
        const epoch = toEpochMs(s, { order, dateOnly: true });
        if (!Number.isNaN(epoch)) {
          matches++;
          if (hasTimeComponent(s)) anyTime = true;
        } else if (casualties.length < 5) {
          casualties.push(rawLabel(v));
        }
      }

      const confidence = matches / nonNullCount;
      if (confidence >= THRESHOLD) {
        return {
          type: anyTime ? 'datetime' : 'date',
          role: 'temporal',
          confidence,
          dateOrder: dateOrder ?? 'ambiguous',
          isPercent: false,
          nullCount,
          distinctCount,
          casualties,
          casualtyCount: nonNullCount - matches,
        };
      }
    } else {
      // Conflict short-circuits straight to text, carrying the dateOrder.
      return {
        type: 'text',
        role: 'ignored',
        confidence: 0,
        dateOrder: 'conflict',
        isPercent: false,
        nullCount,
        distinctCount,
        casualties: [],
        casualtyCount: 0,
      };
    }
  }

  // --- numeric ---------------------------------------------------------
  let numericVerdict = null;
  {
    const casualties = [];
    let matches = 0;
    let percentMatches = 0;
    for (const v of nonNull) {
      const parsed = parseNumberLike(v);
      if (parsed.ok) {
        matches++;
        if (parsed.isPercent) percentMatches++;
      } else if (casualties.length < 5) {
        casualties.push(rawLabel(v));
      }
    }
    const confidence = matches / nonNullCount;
    if (confidence >= THRESHOLD) {
      numericVerdict = {
        type: 'numeric',
        role: 'measure',
        confidence,
        dateOrder: null,
        isPercent: matches > 0 && percentMatches === matches,
        nullCount,
        distinctCount,
        casualties,
        casualtyCount: nonNullCount - matches,
      };
    }
  }

  let verdict;
  if (numericVerdict) {
    verdict = numericVerdict;
  } else {
    // --- categorical vs text -------------------------------------------
    const isCategorical = distinctCount <= 50 || distinctCount / nonNullCount < 0.05;
    verdict = isCategorical
      ? {
          type: 'categorical',
          role: 'dimension',
          confidence: 1,
          dateOrder: null,
          isPercent: false,
          nullCount,
          distinctCount,
          casualties: [],
          casualtyCount: 0,
        }
      : {
          type: 'text',
          role: 'ignored',
          confidence: 1,
          dateOrder: null,
          isPercent: false,
          nullCount,
          distinctCount,
          casualties: [],
          casualtyCount: 0,
        };
  }

  // --- identifier override ---------------------------------------------
  // Runs last, only over numeric/text/categorical verdicts. Fires on:
  //  (a) a consecutive integer sequence (step of exactly 1), regardless
  //      of column name; or
  //  (b) distinct ratio > 0.95 AND the column name matches
  //      /id|no|code|ref|serial/i.
  // (Binding ruling amending the brief's literal "monotonic" wording --
  // see task-2-brief.md.)
  if (verdict.type === 'numeric' || verdict.type === 'text' || verdict.type === 'categorical') {
    const distinctRatio = distinctCount / nonNullCount;
    const consecutive = isConsecutiveIntegerSequence(nonNull);
    const nameMatches = IDENTIFIER_NAME_RE.test(columnName ?? '');

    if (consecutive || (distinctRatio > 0.95 && nameMatches)) {
      verdict = {
        ...verdict,
        type: 'identifier',
        role: 'ignored',
        casualties: [],
        casualtyCount: 0,
      };
    }
  }

  return verdict;
}

function rawLabel(value) {
  if (value instanceof Date) return String(value);
  return String(value ?? '');
}
