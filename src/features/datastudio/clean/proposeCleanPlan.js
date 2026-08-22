// Profile + grid -> the suggested cleaning checklist (spec §8.1).
//
// This module proposes; it never applies. Every step carries a real
// count and a preview of what it will do, because the user is being
// asked to approve a change to their data and "Clean up columns" tells
// them nothing about what they are agreeing to.
//
// `enabled` starts true only for `confidence: 'high'` steps -- the
// mechanical, reversible ones. Anything that drops rows or guesses at
// intent is offered unticked. That split IS the safety model: a step
// nobody read still runs if it is pre-ticked.

import { categoryKey, clusterCategories } from './cleanOps.js';
import { isNullish } from '../profile/inferType.js';

// Steps are emitted in this order so that each op sees the output of the
// ones before it. Trimming has to happen before clustering, or the
// cluster it would have merged is one the trim was about to fix; column
// ops come last so they see the final values.
const OP_ORDER = [
  'trimWhitespace',
  'normalizeNulls',
  'mergeCategories',
  'castType',
  'dropEmptyColumns',
  'dropEmptyRows',
  'dedupeRows',
];

// How many examples a preview names before it starts summarising. Three
// is enough to show the shape of the change without turning a checklist
// row into a paragraph.
const PREVIEW_EXAMPLES = 2;

function collapse(value) {
  return String(value ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/\s+/g, ' ');
}

function columnValues(grid, index) {
  return grid.rows.map((row) => row?.[index]);
}

function isEmptyCell(value) {
  return value === null || value === undefined || isNullish(value);
}

function quote(value) {
  return `"${value}"`;
}

function step(op, column, params, confidence, affectedCount, preview) {
  return {
    // Unique by construction: at most one step per (op, column), and
    // whole-grid ops have no column to collide over.
    id: `${op}:${column ?? '*'}`,
    column,
    op,
    params,
    confidence,
    affectedCount,
    preview,
    enabled: confidence === 'high',
  };
}

// --- per-column proposals ---------------------------------------------

function proposeTrim(values, name) {
  // A Set, not an array: a column of 300 identically-padded values would
  // otherwise preview as `" HR " -> "HR"; " HR " -> "HR"`, which spends
  // the whole example budget saying one thing twice.
  const examples = new Set();
  let affected = 0;

  for (const v of values) {
    if (typeof v !== 'string') continue;
    const cleaned = collapse(v);
    if (cleaned === v) continue;
    affected++;
    if (examples.size < PREVIEW_EXAMPLES) {
      examples.add(`${quote(v)} -> ${quote(cleaned)}`);
    }
  }

  if (affected === 0) return null;
  return step(
    'trimWhitespace', name, {}, 'high', affected,
    `Trim ${affected} value${affected === 1 ? '' : 's'}, e.g. ${[...examples].join('; ')}`,
  );
}

function proposeNulls(values, name) {
  const examples = new Set();
  let affected = 0;

  for (const v of values) {
    // Already null: nothing to normalise, and counting it would inflate
    // the number the user is being asked to approve.
    if (v === null || v === undefined) continue;
    if (!isNullish(v)) continue;
    affected++;
    if (examples.size < PREVIEW_EXAMPLES) examples.add(quote(String(v)));
  }

  if (affected === 0) return null;
  return step(
    'normalizeNulls', name, {}, 'high', affected,
    `Read ${affected} placeholder${affected === 1 ? '' : 's'} as empty, e.g. ${[...examples].join(', ')}`,
  );
}

// Merge proposals come from `clusterCategories`, filtered to the
// clusters worth acting on: a cluster with a single spelling is a
// category, not a merge.
function proposeMerge(values, name) {
  const clusters = clusterCategories(values).filter((c) => c.variants.length > 1);
  if (clusters.length === 0) return null;

  const map = {};
  const examples = [];
  let affected = 0;
  // High unless some cluster needed punctuation stripped to unify. Case
  // and whitespace are typing noise; punctuation might be meaningful,
  // so "PMW-SS" and "PMW SS" get offered rather than assumed.
  let confidence = 'high';

  for (const cluster of clusters) {
    map[cluster.key] = cluster.canonical;

    const collapsedVariants = new Set(cluster.variants.map((v) => collapse(v).toLowerCase()));
    if (collapsedVariants.size > 1) confidence = 'medium';

    for (const variant of cluster.variants) {
      if (variant === cluster.canonical) continue;
      if (examples.length < PREVIEW_EXAMPLES) {
        examples.push(`${quote(variant)} -> ${quote(cluster.canonical)}`);
      }
    }
  }

  for (const v of values) {
    if (isNullish(v)) continue;
    const original = typeof v === 'string' ? v.trim() : String(v);
    const canonical = map[categoryKey(original)];
    if (canonical !== undefined && canonical !== original) affected++;
  }

  if (affected === 0) return null;

  const spellings = clusters.reduce((n, c) => n + c.variants.length, 0);
  return step(
    'mergeCategories', name, { map }, confidence, affected,
    `Merge ${spellings} spellings into ${clusters.length} categor${clusters.length === 1 ? 'y' : 'ies'}: ${examples.join('; ')}`,
  );
}

// Whether a value is already the JS type its column was inferred to be.
// Ruling F6: a cast is proposed only when something actually needs
// coercing, or an already-typed column collects a no-op checklist row
// that the user has to read and dismiss.
function alreadyTyped(value, type, column) {
  if (type === 'numeric') return typeof value === 'number';
  if (type === 'boolean') return typeof value === 'boolean';
  if (type === 'date' || type === 'datetime') return value instanceof Date;
  if (type === 'multi') {
    // Only worth proposing when normalising would actually change the
    // cell: the trailing separator every form export leaves behind, or
    // spaces around the options.
    const separator = column?.separator ?? ';';
    const options = String(value).split(separator).map((p) => p.trim()).filter(Boolean);
    return options.join(separator) === String(value);
  }
  return true;
}

function proposeCast(values, column) {
  const { name, type, dateOrder } = column;
  if (!['numeric', 'boolean', 'date', 'datetime', 'multi'].includes(type)) return null;

  let affected = 0;
  const examples = new Set();

  for (const v of values) {
    if (isEmptyCell(v)) continue;
    if (alreadyTyped(v, type, column)) continue;
    affected++;
    if (examples.size < PREVIEW_EXAMPLES) examples.add(quote(String(v)));
  }

  if (affected === 0) return null;

  const params = { type };
  if (type === 'multi') params.separator = column.separator ?? ';';
  if (type === 'date' || type === 'datetime') {
    // A conflicting order is exactly the case the user has to resolve,
    // so never bake a guess into the plan; fall back to the profile's
    // decision otherwise.
    params.order = dateOrder && dateOrder !== 'conflict' ? dateOrder : 'dmy';
    params.dateOnly = type === 'date';
  }

  return step(
    'castType', name, params, 'high', affected,
    `Read ${affected} value${affected === 1 ? '' : 's'} in ${quote(name)} as ${type}, e.g. ${[...examples].join(', ')}`,
  );
}

// --- whole-grid proposals ---------------------------------------------

function proposeDropColumns(grid) {
  const empty = grid.headers.filter(
    (_, c) => grid.rows.every((row) => isEmptyCell(row?.[c])),
  );
  if (empty.length === 0) return null;

  return step(
    'dropEmptyColumns', null, {}, 'high', empty.length,
    `Drop ${empty.length} empty column${empty.length === 1 ? '' : 's'}: ${empty.map(quote).join(', ')}`,
  );
}

function proposeDropRows(grid) {
  const empty = grid.rows.filter((row) => (row ?? []).every(isEmptyCell)).length;
  if (empty === 0) return null;

  return step(
    'dropEmptyRows', null, {}, 'high', empty,
    `Drop ${empty} row${empty === 1 ? '' : 's'} with nothing in them`,
  );
}

function proposeDedupe(grid) {
  const seen = new Set();
  let duplicates = 0;

  for (const row of grid.rows) {
    const key = JSON.stringify((row ?? []).map(
      (cell) => (cell instanceof Date ? `D${cell.getTime()}` : cell),
    ));
    if (seen.has(key)) duplicates++;
    else seen.add(key);
  }

  if (duplicates === 0) return null;

  // Medium, never high: two identical rows can be two real events that
  // happen to record the same values. Only the user knows which.
  return step(
    'dedupeRows', null, {}, 'medium', duplicates,
    `Remove ${duplicates} duplicate row${duplicates === 1 ? '' : 's'}, keeping the first of each`,
  );
}

export function proposeCleanPlan(profile, grid) {
  const steps = [];

  for (const column of profile.columns) {
    const values = columnValues(grid, column.index);

    steps.push(proposeTrim(values, column.name));
    steps.push(proposeNulls(values, column.name));
    // Clustering is only meaningful for columns that are categories.
    // Running it over free text with a distinct value per row would
    // produce one cluster per row and propose nothing, at the cost of a
    // full pass.
    if (column.role === 'dimension') steps.push(proposeMerge(values, column.name));
    steps.push(proposeCast(values, column));
  }

  steps.push(proposeDropColumns(grid));
  steps.push(proposeDropRows(grid));
  steps.push(proposeDedupe(grid));

  return steps
    .filter(Boolean)
    .sort((a, b) => OP_ORDER.indexOf(a.op) - OP_ORDER.indexOf(b.op));
}
