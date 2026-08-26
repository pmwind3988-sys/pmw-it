import { normaliseCode } from './identity.js';
import { TRACKED } from './assetKinds.js';
import { unitsOf, countPerItem, UNIT_FIELDS } from './units.js';
import { needsDetails } from './detailsPending.js';

const UNIT_KEYS = UNIT_FIELDS.map((field) => field.key);

/**
 * Whether every item on the row wears a sticker. A bulk line's labels are on
 * its items, so a box of two with one labelled is not a labelled row.
 */
function labelled(asset) {
  if (asset?.trackingMode === TRACKED) return Boolean(String(asset?.assetTag ?? '').trim());

  const items = unitsOf(asset);
  return items.every((unit) => unit.assetTag);
}

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
  'macAddress', 'assetTag', 'location', 'supplier', 'poNumber', 'doNumber',
  'batchTitle', 'remarks', 'specSummary', 'assignedTo',
];

export function haystack(asset) {
  const parts = HAYSTACK_FIELDS.map((field) => asset?.[field]).filter(Boolean);

  // A bulk line's serials and labels live on its individual items, so a search
  // that only read the row would fail to find a box of tabs by the serial of
  // the tab in somebody's hand — the one search anybody actually runs.
  for (const unit of unitsOf(asset)) {
    for (const key of UNIT_KEYS) if (unit[key]) parts.push(unit[key]);
  }

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
 * unlabelled, pending }`. An absent or empty value is "no opinion", never "match
 * nothing" — the filter bar starts empty and must show everything.
 */
export function filterAssets(assets = [], filters = {}) {
  return assets.filter((asset) => {
    if (filters.category && asset.category !== filters.category) return false;
    // Status and condition are per item on a bulk line, so a row is shown when
    // ANY of the things on it is in that state. Filtering on the row alone
    // would hide the one faulty tab behind a line that has no condition.
    if (filters.status && !countPerItem(asset, 'status', filters.status, 'In stock')) return false;
    if (filters.condition && !countPerItem(asset, 'condition', filters.condition)) return false;
    if (filters.location && asset.location !== filters.location) return false;
    if (filters.trackingMode && asset.trackingMode !== filters.trackingMode) return false;
    if (filters.unlabelled && labelled(asset)) return false;
    if (filters.pending && !needsDetails(asset)) return false;
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
