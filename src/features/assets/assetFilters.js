import { normaliseCode } from './identity.js';

/**
 * Finding one thing in the register.
 *
 * Search reaches every code on the row — serial, part number, MAC, sticker
 * label and the codes nobody could place — because the thing somebody is
 * holding when they search is a barcode, and it does not matter which field it
 * turned out to belong in.
 */

const HAYSTACK_FIELDS = [
  'title', 'manufacturer', 'model', 'serialNumber', 'partNumber',
  'macAddress', 'assetTag', 'location', 'supplier', 'poNumber',
  'batchTitle', 'remarks', 'specSummary', 'assignedTo',
];

export function haystack(asset) {
  const parts = HAYSTACK_FIELDS.map((field) => asset?.[field]).filter(Boolean);
  return [...parts, ...(asset?.additionalCodes ?? [])].join(' ').toLowerCase();
}

/**
 * A scanned code and a typed one have to find the same row: the label reads
 * `CN0ABC123`, somebody types `cn0abc 123`, and both must match. So a query
 * that looks like a code is also tried with its spacing stripped.
 */
export function matchesQuery(asset, query) {
  const term = String(query ?? '').trim().toLowerCase();
  if (!term) return true;

  const text = haystack(asset);
  if (text.includes(term)) return true;

  const code = normaliseCode(term).toLowerCase();
  if (code === term || !code) return false;

  return text.replace(/\s+/g, '').includes(code);
}

/**
 * `filters` is `{ query, category, status, condition, location, trackingMode,
 * unlabelled }`. An absent or empty value is "no opinion", never "match
 * nothing" — the filter bar starts empty and must show everything.
 */
export function filterAssets(assets = [], filters = {}) {
  return assets.filter((asset) => {
    if (filters.category && asset.category !== filters.category) return false;
    if (filters.status && (asset.status || 'In stock') !== filters.status) return false;
    if (filters.condition && asset.condition !== filters.condition) return false;
    if (filters.location && asset.location !== filters.location) return false;
    if (filters.trackingMode && asset.trackingMode !== filters.trackingMode) return false;
    if (filters.unlabelled && String(asset.assetTag ?? '').trim()) return false;
    return matchesQuery(asset, filters.query);
  });
}

const SORTERS = {
  name: (a, b) => String(a.title ?? '').localeCompare(String(b.title ?? '')),
  category: (a, b) => String(a.category ?? '').localeCompare(String(b.category ?? '')),
  quantity: (a, b) => (b.quantity ?? 0) - (a.quantity ?? 0),
  // Newest first, and rows with no date sink rather than leading — an item
  // saved before the column existed is not the newest thing in the register.
  arrived: (a, b) => (b.arrivedOn ?? 0) - (a.arrivedOn ?? 0),
};

export function sortAssets(assets, key = 'arrived') {
  return [...assets].sort(SORTERS[key] ?? SORTERS.arrived);
}

/** The values actually present, so a filter can never offer an empty result. */
export function optionsFor(assets = [], field) {
  const seen = new Set();
  for (const asset of assets) {
    const value = asset?.[field];
    if (value) seen.add(String(value));
  }
  return [...seen].sort((a, b) => a.localeCompare(b));
}
