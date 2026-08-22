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
// Non-breaking space (U+00A0) needs no separate handling: JS's `\s` class
// already includes it, so the generic whitespace stripping below (and
// `String.prototype.trim()`, at the edges) covers it for free. Written as
// escape sequences rather than literal characters on purpose: a literal
// invisible character in source does not survive being retyped or diffed
// and can silently rot into a no-op.
const ZERO_WIDTH_RE = /[\u200B-\u200D\uFEFF]/g;

// Normalises a raw cell value to a trimmed string for comparison/parsing.
// `Date` objects are the one exception -- callers must check for those
// before calling this, since stringifying a Date here would destroy it.
function normalizeToString(value) {
  const s = String(value ?? '');
  return s.replace(ZERO_WIDTH_RE, '').trim();
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
  const rawTrimmed = raw.replace(ZERO_WIDTH_RE, '').trim();
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
    // A leading '-' on top of an already-detected accounting negative is a
    // double negation ("(-5)" means -(-5) = 5), so this must toggle, not
    // just force `true`.
    negative = !negative;
    s = s.slice(1).trim();
  } else if (/^\+/.test(s)) {
    s = s.slice(1).trim();
  }

  // Strip thousands separators (commas) and any remaining whitespace used
  // as a separator.
  s = s.replace(/,/g, '').replace(/\s+/g, '');

  // The leading-zero guard has to run a second time. The first pass sees
  // the raw string, but '$007', '(007)' and '-007' only expose their
  // leading zero once the currency symbol, parentheses and sign are gone.
  // Without this, a cost centre written as '$007' parses as 7 -- exactly
  // the silent mangling spec section 7.4 exists to prevent.
  if (LEADING_ZERO_RE.test(s)) return fail;

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
// A run this short is not evidence of anything: ['10','11'] is two
// ordinary measurements as readily as it is two row numbers, and preview
// slices and tiny sheets would otherwise produce spurious identifier
// verdicts that demote a real measure to role 'ignored'.
export const MIN_IDENTIFIER_RUN = 5;

function isConsecutiveIntegerSequence(rawValues) {
  if (rawValues.length < MIN_IDENTIFIER_RUN) return false;

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

const IDENTIFIER_NAME_TOKEN_RE = /^(ids?|nos?|codes?|refs?|serials?)$/i;

// Splits a header into words, breaking on separators AND on camelCase
// boundaries, so 'EmpID' and 'Emp_ID' tokenise the same way as 'Emp ID'.
function nameTokens(name) {
  return String(name ?? '')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
}

// Whether a header names an identifier. Matching is per-token, not by
// substring: the brief's literal /id|no|code|ref|serial/i is unanchored,
// so 'Paid Amount' (contains 'id'), 'Width', 'Notes' and 'Income' all
// matched it, and a genuine measure with distinct values was silently
// demoted to role 'ignored' -- the same class of harm as misclassifying
// a date column. Token matching keeps every intended hit ('Employee ID',
// 'Ref No', 'Serial Number', 'Cost Code') and drops the collisions.
function matchesIdentifierName(name) {
  return nameTokens(name).some((token) => IDENTIFIER_NAME_TOKEN_RE.test(token));
}

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

export const MULTI_SEPARATORS = [';', '|'];

// A multi-select answer is many options in one cell. Three conditions
// have to hold together, and each one rules out a different impostor:
//
//   * most values carry the separator          -- or it is prose that
//                                                 happens to contain one
//   * the average cell holds more than one     -- or it is a plain
//     option                                      category with a stray
//                                                 trailing separator
//   * the options repeat across rows           -- or the "options" are
//                                                 sentences, and every
//                                                 one is unique
//
// The last is the load-bearing one. Free text split on ';' produces a
// distinct part for almost every row; a real multi-select reuses a small
// fixed menu.
const MULTI_MIN_SEPARATED_RATIO = 0.6;
const MULTI_MIN_PARTS_PER_VALUE = 1.2;
const MULTI_MAX_DISTINCT_RATIO = 0.5;
const MULTI_MAX_DISTINCT_PARTS = 60;

function detectMultiSeparator(nonNull) {
  const strings = nonNull.filter((v) => !(v instanceof Date)).map(normalizeToString);
  if (strings.length === 0) return null;

  for (const separator of MULTI_SEPARATORS) {
    let separated = 0;
    let parts = 0;
    const distinct = new Set();

    for (const s of strings) {
      if (s.includes(separator)) separated++;
      for (const part of s.split(separator)) {
        const trimmed = part.trim();
        if (trimmed === '') continue;
        parts++;
        distinct.add(trimmed.toLowerCase());
      }
    }

    if (parts === 0) continue;
    if (separated / strings.length < MULTI_MIN_SEPARATED_RATIO) continue;
    if (parts / strings.length < MULTI_MIN_PARTS_PER_VALUE) continue;
    if (distinct.size > MULTI_MAX_DISTINCT_PARTS) continue;
    if (distinct.size / parts > MULTI_MAX_DISTINCT_RATIO) continue;

    return separator;
  }

  return null;
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
  //
  // 'conflict' means two values in the column disagree on day/month
  // placement -- detectDateOrder found e.g. both a value only valid as
  // dmy and a value only valid as mdy. That does NOT mean the column
  // isn't full of dates; it means we cannot commit to a single order for
  // it. The rule (binding ruling amending the brief): a conflict blocks
  // the column ONLY when the datetime candidate would otherwise have won
  // the cascade. If it would not have won anyway (e.g. two stray date
  // strings in an otherwise-numeric column), the conflict is irrelevant
  // information carried onto whatever verdict does win, not a reason to
  // short-circuit the whole cascade.
  const stringValues = nonNull.filter((v) => !(v instanceof Date));
  const dateOrder = stringValues.length > 0 ? detectDateOrder(stringValues.map(normalizeToString)) : null;
  const isConflict = dateOrder === 'conflict';

  {
    const order = isConflict ? null : (dateOrder === null ? 'dmy' : dateOrder);
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
      let epoch;
      if (isConflict) {
        // The order is genuinely unresolved for this column, so this is a
        // structural check only: does the value parse as a date under
        // EITHER candidate order? The resulting epoch (and thus which
        // order "wins") is never used -- a conflicting column can never
        // commit to one order -- only whether the value looks like a date
        // at all, to decide whether datetime would have won the cascade.
        const dmyEpoch = toEpochMs(s, { order: 'dmy', dateOnly: true });
        epoch = !Number.isNaN(dmyEpoch) ? dmyEpoch : toEpochMs(s, { order: 'mdy', dateOnly: true });
      } else {
        epoch = toEpochMs(s, { order, dateOnly: true });
      }
      if (!Number.isNaN(epoch)) {
        matches++;
        if (hasTimeComponent(s)) anyTime = true;
      } else if (casualties.length < 5) {
        casualties.push(rawLabel(v));
      }
    }

    const confidence = matches / nonNullCount;
    if (confidence >= THRESHOLD) {
      if (isConflict) {
        // The datetime candidate would have won, but the order cannot be
        // determined -- leave it as text for the user to resolve rather
        // than silently guessing.
        return {
          type: 'text',
          role: 'ignored',
          confidence: 1,
          dateOrder: 'conflict',
          isPercent: false,
          nullCount,
          distinctCount,
          casualties: [],
          casualtyCount: 0,
        };
      }
      return {
        type: anyTime ? 'datetime' : 'date',
        role: 'temporal',
        confidence,
        dateOrder,
        isPercent: false,
        nullCount,
        distinctCount,
        casualties,
        casualtyCount: nonNullCount - matches,
      };
    }
    // Datetime candidate declined (below threshold). If this was a
    // conflict, it's now irrelevant to the type decision -- fall through
    // to numeric/categorical/text/identifier, and tag `dateOrder:
    // 'conflict'` onto the eventual verdict purely as information.
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
    // --- multi-select ----------------------------------------------------
    // Ahead of the categorical/text split: a multi-select column is
    // categorical-looking (few distinct cells) or text-looking (many)
    // depending only on how many combinations people happened to pick,
    // so neither verdict can be trusted to fall out correctly on its own.
    const separator = detectMultiSeparator(nonNull);

    if (separator) {
      verdict = {
        type: 'multi',
        role: 'dimension',
        confidence: 1,
        dateOrder: null,
        isPercent: false,
        separator,
        nullCount,
        distinctCount,
        casualties: [],
        casualtyCount: 0,
      };
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
  }

  // --- identifier override ---------------------------------------------
  // Runs last, only over numeric/text/categorical verdicts. Fires on:
  //  (a) a consecutive integer sequence (step of exactly 1), regardless
  //      of column name; or
  //  (b) distinct ratio > 0.95 AND the column name carries an
  //      identifier word as a whole token (see `matchesIdentifierName`).
  // (Binding ruling amending the brief's literal "monotonic" wording --
  // see task-2-brief.md.)
  if (verdict.type === 'numeric' || verdict.type === 'text' || verdict.type === 'categorical') {
    const distinctRatio = distinctCount / nonNullCount;
    const consecutive = isConsecutiveIntegerSequence(nonNull);
    const nameMatches = matchesIdentifierName(columnName);

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

  if (isConflict) {
    // Informational only at this point -- the datetime candidate did not
    // win, so the conflict didn't block anything, but the user should
    // still be told this column had a date-order disagreement.
    verdict = { ...verdict, dateOrder: 'conflict' };
  }

  return verdict;
}

function rawLabel(value) {
  if (value instanceof Date) return String(value);
  return String(value ?? '');
}
