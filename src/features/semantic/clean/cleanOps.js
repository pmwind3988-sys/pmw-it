// The cleaning operations -- spec §8.2.
//
// Ten pure functions. Seven take a column's values and return new
// values; three take the whole grid and return a new grid. None of them
// mutates its input, because the clean plan is a checklist the user
// toggles: a step that has been unticked has to leave no trace, and the
// only way to guarantee that is to never write over the source.
//
// The recurring decision in this file is what to do with a value that
// does not fit. The answer is always `null`, never a guess. A cost of
// "pending" becomes null and shows up as a gap; coerced to 0 it would
// quietly drag every average in the dashboard down and look like data.

import {
  NULL_TOKENS, isNullish, parseNumberLike, parseBooleanLike,
} from '../profile/inferType.js';
import { toEpochMs } from '../../../utils/malaysiaTime.js';

// ---------------------------------------------------------------------
// Column operations
// ---------------------------------------------------------------------

// Trims the ends, collapses internal whitespace runs to one space, and
// removes zero-width characters. Non-strings (numbers, Date objects) are
// passed through untouched -- stringifying a Date to trim it would
// destroy it.
export function trimWhitespace(values) {
  return values.map((v) => {
    if (typeof v !== 'string') return v;
    return v
      // Zero-width characters (ZWSP/ZWNJ/ZWJ/BOM) are NOT in JS's `\s`,
      // so nothing else here removes them. Escape sequences, never
      // literals: an invisible character does not survive being retyped.
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .trim()
      // `\s` covers the non-breaking space (U+00A0) as well as ordinary
      // ones, so a value padded with NBSPs by a copy-paste out of a web
      // page collapses the same way.
      .replace(/\s+/g, ' ');
  });
}

// Every spreadsheet dialect for "nothing here" -- '', '-', 'N/A', 'NIL',
// and Excel's error text -- becomes an actual null, so downstream code
// has one empty value to check rather than a dozen.
export function normalizeNulls(values) {
  return values.map((v) => (isNullish(v) ? null : v));
}

export function parseNumber(values) {
  return values.map((v) => {
    if (isNullish(v)) return null;
    const parsed = parseNumberLike(v);
    return parsed.ok ? parsed.value : null;
  });
}

// Rewrites case variants onto the spelling that appears most often, so
// 'HR', 'hr' and 'Hr' stop being three categories on a chart. Ties go to
// whichever spelling appeared first, which makes the result stable
// rather than dependent on iteration order.
export function unifyCase(values) {
  const spellings = new Map();

  for (const v of values) {
    if (typeof v !== 'string' || isNullish(v)) continue;
    const key = v.trim().toLowerCase();
    let bySpelling = spellings.get(key);
    if (!bySpelling) {
      bySpelling = new Map();
      spellings.set(key, bySpelling);
    }
    const trimmed = v.trim();
    bySpelling.set(trimmed, (bySpelling.get(trimmed) ?? 0) + 1);
  }

  const canonical = new Map();
  for (const [key, bySpelling] of spellings) {
    let best = null;
    let bestCount = -1;
    for (const [spelling, count] of bySpelling) {
      if (count > bestCount) {
        best = spelling;
        bestCount = count;
      }
    }
    canonical.set(key, best);
  }

  return values.map((v) => {
    if (typeof v !== 'string' || isNullish(v)) return v;
    return canonical.get(v.trim().toLowerCase()) ?? v;
  });
}

// The key two spellings must share to count as the same category:
// lowercase, trimmed, internal runs collapsed, punctuation stripped.
//
// This is deliberately the WHOLE of the matching rule. Fuzzy or
// edit-distance matching is excluded by spec §3: 'Dept A' and 'Dept B'
// are edit distance 1, so a fuzzy merge silently destroys the
// distinction between two real departments. Demos beautifully, corrupts
// data quietly.
export function categoryKey(value) {
  return String(value ?? '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Groups values by `categoryKey`, naming each group after its most
// frequent original spelling.
//
// Every group is returned, including groups of one. The merge proposal
// step is what filters down to groups worth showing the user (more than
// one spelling); returning only those from here would mean a caller
// asking "what categories are in this column?" got a misleading answer,
// and it is the same question.
export function clusterCategories(values) {
  const groups = new Map();

  for (const v of values) {
    if (isNullish(v)) continue;
    const original = typeof v === 'string' ? v.trim() : String(v);
    const key = categoryKey(original);
    if (key === '') continue;

    let group = groups.get(key);
    if (!group) {
      group = { key, count: 0, spellings: new Map(), order: groups.size };
      groups.set(key, group);
    }
    group.count++;
    group.spellings.set(original, (group.spellings.get(original) ?? 0) + 1);
  }

  return [...groups.values()]
    .map((group) => {
      let canonical = null;
      let bestCount = -1;
      for (const [spelling, count] of group.spellings) {
        if (count > bestCount) {
          canonical = spelling;
          bestCount = count;
        }
      }
      return {
        key: group.key,
        canonical,
        variants: [...group.spellings.keys()],
        count: group.count,
        order: group.order,
      };
    })
    // Biggest group first, so the merges worth a user's attention are at
    // the top of the list. Ties keep first-seen order for stability.
    .sort((a, b) => b.count - a.count || a.order - b.order)
    .map(({ key, canonical, variants, count }) => ({ key, canonical, variants, count }));
}

// Rewrites values to a canonical spelling using a map keyed by
// `categoryKey`. Values whose key is absent from the map are left
// exactly as they were -- this op only ever applies merges the user has
// actually agreed to.
export function mergeCategories(values, params = {}) {
  const map = params.map ?? {};
  return values.map((v) => {
    if (isNullish(v)) return v;
    const key = categoryKey(v);
    return Object.prototype.hasOwnProperty.call(map, key) ? map[key] : v;
  });
}

// Dates to epoch milliseconds. Anything that does not parse becomes
// null rather than an Invalid Date, which would otherwise travel all the
// way to a time axis before failing.
export function parseDate(values, params = {}) {
  const { order = 'dmy', sourceZone = 'local', dateOnly = false } = params;
  return values.map((v) => {
    if (isNullish(v)) return null;
    const ms = toEpochMs(v, { order, sourceZone, dateOnly });
    return Number.isNaN(ms) ? null : ms;
  });
}

// Forces a column to a type the user picked, coercing what fits and
// nulling what does not.
export function castType(values, params = {}) {
  const { type } = params;

  switch (type) {
    case 'numeric':
      return parseNumber(values);
    case 'boolean':
      return values.map((v) => {
        if (isNullish(v)) return null;
        const parsed = parseBooleanLike(v);
        return parsed.ok ? parsed.value : null;
      });
    case 'date':
      return parseDate(values, { ...params, dateOnly: true });
    case 'datetime':
      return parseDate(values, params);
    case 'multi': {
      const separator = params.separator ?? ';';
      return values.map((v) => {
        if (isNullish(v)) return null;
        const options = String(v)
          .split(separator)
          .map((part) => part.trim())
          .filter(Boolean);
        // A cell of nothing but separators held no options at all, so it
        // is empty -- not an option named "".
        return options.length > 0 ? options.join(separator) : null;
      });
    }
    case 'text':
    case 'categorical':
      return values.map((v) => {
        if (isNullish(v)) return null;
        return typeof v === 'string' ? v.trim() : String(v);
      });
    default:
      // An unknown type is a bug in the caller, not a licence to mangle
      // the column. Hand it back untouched.
      return values.slice();
  }
}

// ---------------------------------------------------------------------
// Whole-grid operations
// ---------------------------------------------------------------------

function isEmptyCell(value) {
  return value === null || value === undefined || isNullish(value);
}

export function dropEmptyColumns(grid) {
  const { headers, rows } = grid;
  const keep = headers.map((_, c) => rows.some((row) => !isEmptyCell(row?.[c])));

  return {
    headers: headers.filter((_, c) => keep[c]),
    rows: rows.map((row) => headers.map((_, c) => row?.[c]).filter((_, c) => keep[c])),
  };
}

export function dropEmptyRows(grid) {
  return {
    headers: grid.headers,
    rows: grid.rows.filter((row) => (row ?? []).some((cell) => !isEmptyCell(cell))),
  };
}

// Exact duplicate rows only -- same values in the same columns. The
// first occurrence is kept, so the order the user's sheet was in
// survives.
export function dedupeRows(grid) {
  const seen = new Set();
  const rows = [];

  for (const row of grid.rows) {
    // JSON is a sound identity here because every cell at this point is
    // a primitive or a Date, both of which stringify unambiguously.
    const key = JSON.stringify((row ?? []).map(
      (cell) => (cell instanceof Date ? `D${cell.getTime()}` : cell),
    ));
    if (seen.has(key)) continue;
    seen.add(key);
    rows.push(row);
  }

  return { headers: grid.headers, rows };
}

// Re-exported so the clean layer has one import surface and callers do
// not reach back into the profiling module for a constant.
export { NULL_TOKENS };
