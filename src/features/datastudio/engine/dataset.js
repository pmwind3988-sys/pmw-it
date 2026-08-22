// The columnar store every later phase reads -- spec §6.1.
//
// Charts ask column-shaped questions ("sum Amount grouped by Dept"), so
// the data is stored column-shaped: one typed array per column instead
// of one object per row. At 100k rows that is the difference between a
// few megabytes of contiguous numbers and 100k objects for the garbage
// collector to chase, and a filter pass becomes a tight loop over a
// Float64Array rather than a property lookup per row.
//
// NULL ENCODINGS ARE A CONTRACT (spec §6.1). Every consumer -- filter
// masks, aggregation, chart tiles -- tests for these exact values:
//
//   numeric, date, datetime  ->  NaN
//   categorical              ->  -1   (no dictionary entry has this code)
//   boolean                  ->   2   (so true=1 and false=0 stay usable)
//   text, identifier, empty  ->  null
//
// The rule behind all four: a missing value must never collide with a
// real one. Encoding a missing number as 0 makes it a real zero that
// drags every average down; encoding a missing category as 0 makes it
// the first real category.

export const NULL_NUMBER = NaN;
export const NULL_CODE = -1;
export const NULL_BOOL = 2;

const NUMERIC_TYPES = new Set(['numeric']);
const TEMPORAL_TYPES = new Set(['date', 'datetime']);
const CATEGORICAL_TYPES = new Set(['categorical']);
const MULTI_TYPES = new Set(['multi']);
const BOOLEAN_TYPES = new Set(['boolean']);

function isMissing(value) {
  return value === null || value === undefined
    || (typeof value === 'number' && Number.isNaN(value));
}

function encodeNumeric(values) {
  const out = new Float64Array(values.length);
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (isMissing(v)) {
      out[i] = NULL_NUMBER;
      continue;
    }
    const n = typeof v === 'number' ? v : Number(v);
    out[i] = Number.isFinite(n) ? n : NULL_NUMBER;
  }
  return out;
}

function encodeTemporal(values) {
  const out = new Float64Array(values.length);
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (isMissing(v)) {
      out[i] = NULL_NUMBER;
      continue;
    }
    // Epoch milliseconds either way -- a Date that survived cleaning
    // untouched and a string the cast already converted must land in the
    // same representation, or a time axis gets both and plots neither.
    const ms = v instanceof Date ? v.getTime() : Number(v);
    out[i] = Number.isFinite(ms) ? ms : NULL_NUMBER;
  }
  return out;
}

// Dictionary encoding: the column becomes small integer codes into a
// list of the distinct labels. Grouping then compares integers rather
// than strings, and the labels are stored once instead of once per row.
// First-appearance order, so the dictionary is stable across runs.
function encodeCategorical(values) {
  const dictionary = [];
  const codes = new Map();
  const out = new Int32Array(values.length);

  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (isMissing(v)) {
      out[i] = NULL_CODE;
      continue;
    }
    const label = typeof v === 'string' ? v : String(v);
    let code = codes.get(label);
    if (code === undefined) {
      code = dictionary.length;
      dictionary.push(label);
      codes.set(label, code);
    }
    out[i] = code;
  }

  return { values: out, dictionary };
}

// Multi-select: one row holds several options, so a flat code array plus
// per-row offsets (compressed-sparse-row) rather than one code per row.
// Row r's options are values[offsets[r] .. offsets[r + 1]). Both arrays
// are typed, so this survives structured clone and IndexedDB untouched
// the same way every other column does.
//
// A row with no options is offsets[r] === offsets[r + 1] -- an empty
// range, which is the null encoding for this type. There is no sentinel
// code, because there is no single slot to put one in.
function encodeMulti(values, separator = ';') {
  const dictionary = [];
  const codes = new Map();
  const flat = [];
  const offsets = new Int32Array(values.length + 1);

  for (let i = 0; i < values.length; i++) {
    offsets[i] = flat.length;
    const v = values[i];
    if (isMissing(v)) continue;
    for (const part of String(v).split(separator)) {
      const label = part.trim();
      if (label === '') continue;
      let code = codes.get(label);
      if (code === undefined) {
        code = dictionary.length;
        dictionary.push(label);
        codes.set(label, code);
      }
      flat.push(code);
    }
  }
  offsets[values.length] = flat.length;

  return { values: Int32Array.from(flat), offsets, dictionary };
}

function encodeBoolean(values) {
  const out = new Uint8Array(values.length);
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (isMissing(v)) {
      out[i] = NULL_BOOL;
      continue;
    }
    if (typeof v === 'boolean') {
      out[i] = v ? 1 : 0;
      continue;
    }
    const s = String(v).trim().toLowerCase();
    if (s === 'true' || s === 'yes' || s === 'y') out[i] = 1;
    else if (s === 'false' || s === 'no' || s === 'n') out[i] = 0;
    else out[i] = NULL_BOOL;
  }
  return out;
}

function encodeText(values) {
  return values.map((v) => (isMissing(v) ? null : String(v)));
}

/**
 * Encodes cleaned column values into the typed columnar store.
 *
 * `columns` is parallel to `headers`: one array of cleaned values each.
 * `profile` supplies the type and role, so the encoding follows the
 * verdict the user saw and approved rather than re-guessing here.
 */
export function buildDataset({ headers, columns, profile }) {
  const byProfileName = new Map((profile?.columns ?? []).map((c) => [c.name, c]));

  const built = headers.map((name, i) => {
    const raw = columns[i] ?? [];
    const meta = byProfileName.get(name) ?? {};
    const type = meta.type ?? 'text';
    const role = meta.role ?? 'ignored';

    let values;
    let dictionary = null;
    let offsets = null;

    if (NUMERIC_TYPES.has(type)) {
      values = encodeNumeric(raw);
    } else if (TEMPORAL_TYPES.has(type)) {
      values = encodeTemporal(raw);
    } else if (MULTI_TYPES.has(type)) {
      ({ values, offsets, dictionary } = encodeMulti(raw, meta.separator ?? ';'));
    } else if (CATEGORICAL_TYPES.has(type)) {
      ({ values, dictionary } = encodeCategorical(raw));
    } else if (BOOLEAN_TYPES.has(type)) {
      values = encodeBoolean(raw);
    } else {
      values = encodeText(raw);
    }

    return {
      name,
      type,
      role,
      values,
      dictionary,
      // Null for every type but `multi`, where it is what turns the flat
      // option array back into rows.
      offsets,
      isPercent: Boolean(meta.isPercent),
      // `dateOnly` tells a time axis whether to render a time of day. A
      // date-only column has no meaningful hours, so showing 00:00
      // against every point would be inventing precision.
      dateOnly: type === 'date',
      sourceZone: meta.sourceZone ?? 'local',
    };
  });

  // Derived from the RAW input, not from the first built column. A multi
  // column's `values` is the flat option array and is longer than the
  // grid, so reading the length off it reports the wrong row count for
  // the whole dataset -- and every mask allocated from it would then be
  // the wrong size.
  const rowCount = headers.length > 0 ? (columns[0]?.length ?? 0) : 0;

  return {
    rowCount,
    columns: built,
    // Name -> position, so a tile spec that names a column can reach it
    // without scanning. Names are unique by construction (`toGrid`
    // de-duplicates them), which is what makes this safe.
    byName: new Map(built.map((c, i) => [c.name, i])),
  };
}

// Convenience for the many callers that hold a name rather than an
// index. Returns undefined for an unknown name rather than throwing:
// a saved dashboard can name a column that this dataset does not have,
// and that is a tile to skip, not a crash.
export function columnByName(dataset, name) {
  const index = dataset.byName.get(name);
  return index === undefined ? undefined : dataset.columns[index];
}
