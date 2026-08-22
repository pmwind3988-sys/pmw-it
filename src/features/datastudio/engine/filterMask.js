// Filters -> one shared row mask -- spec §10.3.
//
// Every tile on the canvas is a pure function of (dataset, mask, spec).
// Rather than each tile filtering the dataset for itself, the filters
// are resolved ONCE into a `Uint8Array` where 1 means "keep this row",
// and every tile reads the same array. Twelve tiles then cost one pass
// over the data instead of twelve.
//
// The rule that matters most in the whole feature lives here: the tile
// that originated a cross-filter selection is NOT filtered by it. Click
// "HR" on a department bar chart and the other tiles narrow to HR --
// but the bar chart itself keeps all its bars, because a chart that
// filtered itself down to the one bar you just clicked would delete the
// context you need to click anything else. That is `maskFor`.

function columnOf(dataset, name) {
  const index = dataset.byName.get(name);
  return index === undefined ? undefined : dataset.columns[index];
}

// Applies one filter by clearing bits in `mask`. Never sets a bit:
// filters only ever narrow, so ANDing them is just running them in
// sequence over the same array.
function applyFilter(dataset, mask, filter) {
  const column = columnOf(dataset, filter?.column);
  // A saved dashboard can name a column this dataset does not have.
  // Ignoring the filter shows more than was asked for; treating it as
  // matching nothing shows an empty dashboard with no explanation.
  // Ignoring is the lesser harm, and the filter bar still lists what is
  // actually in force.
  if (!column) return;

  const { values } = column;

  if (filter.kind === 'range') {
    const min = filter.min ?? -Infinity;
    const max = filter.max ?? Infinity;
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      // NaN is the null encoding for numeric and temporal columns, and
      // every comparison against it is false -- so this also, correctly,
      // drops missing values from a range filter rather than counting
      // them at either end.
      if (!(v >= min && v <= max)) mask[i] = 0;
    }
    return;
  }

  if (filter.kind !== 'in') return;

  const wanted = filter.values ?? [];
  // An empty membership filter means "no constraint chosen yet", not
  // "match nothing" -- a half-built filter must not blank the dashboard.
  if (wanted.length === 0) return;

  if (column.dictionary) {
    // Resolve the labels to integer codes ONCE, then compare integers in
    // the row loop. Comparing strings per row is the difference between
    // a scan and a stall at 100k rows.
    const codes = new Set();
    for (const label of wanted) {
      const code = column.dictionary.indexOf(label);
      if (code !== -1) codes.add(code);
    }
    for (let i = 0; i < values.length; i++) {
      // The null code (-1) is never in `codes`, since it is never a
      // dictionary position, so missing categories drop out here.
      if (!codes.has(values[i])) mask[i] = 0;
    }
    return;
  }

  if (column.type === 'boolean') {
    const wantTrue = wanted.some((v) => v === true || String(v).toLowerCase() === 'true');
    const wantFalse = wanted.some((v) => v === false || String(v).toLowerCase() === 'false');
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      const keep = (v === 1 && wantTrue) || (v === 0 && wantFalse);
      if (!keep) mask[i] = 0;
    }
    return;
  }

  // Text and identifier columns hold plain strings.
  const wantedSet = new Set(wanted.map((v) => String(v)));
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v === null || !wantedSet.has(v)) mask[i] = 0;
  }
}

export function buildMask(dataset, filters) {
  const mask = new Uint8Array(dataset.rowCount).fill(1);
  for (const filter of filters ?? []) applyFilter(dataset, mask, filter);
  return mask;
}

/**
 * The mask one specific tile should read (spec §10.3).
 *
 * `selection` is the current cross-filter click, or null. It applies to
 * every tile EXCEPT the one it came from -- see the note at the top of
 * this file. Global filters, by contrast, apply to everything including
 * the source tile: the user set those deliberately from the filter bar,
 * not by clicking a chart.
 */
export function maskFor(dataset, globalFilters, selection, tileId) {
  const applySelection = Boolean(selection) && selection.sourceTileId !== tileId;
  const filters = applySelection
    ? [...(globalFilters ?? []),
      { column: selection.column, kind: 'in', values: selection.values }]
    : (globalFilters ?? []);
  return buildMask(dataset, filters);
}

// The cache key deliberately does NOT include the tile id -- only
// whether the selection applies to it. Keying on the tile id would give
// every tile its own identical array, which is exactly the duplicated
// work the shared mask exists to avoid; keying on neither would serve
// the source tile the filtered mask and silently undo self-exclusion.
function signature(globalFilters, selection, applySelection) {
  return JSON.stringify({
    globals: globalFilters ?? [],
    selection: applySelection
      ? { column: selection.column, values: selection.values }
      : null,
  });
}

export function createMaskCache() {
  // Keyed by dataset identity, so a new dataset (a re-clean, a new
  // import) cannot be served a mask of the wrong length -- and the old
  // entries become collectable the moment nothing holds that dataset.
  const byDataset = new WeakMap();

  return {
    get(dataset, globalFilters, selection, tileId) {
      let cache = byDataset.get(dataset);
      if (!cache) {
        cache = new Map();
        byDataset.set(dataset, cache);
      }

      const applySelection = Boolean(selection) && selection.sourceTileId !== tileId;
      const key = signature(globalFilters, selection, applySelection);

      const hit = cache.get(key);
      if (hit) return hit;

      const mask = maskFor(dataset, globalFilters, selection, tileId);
      cache.set(key, mask);
      return mask;
    },
  };
}
