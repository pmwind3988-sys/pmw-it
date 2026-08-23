import { TRACKED } from '../assetKinds.js';

/**
 * The figures at the top of the register.
 *
 * Everything counts UNITS, not rows. A row reading "Logitech B100 × 20" is
 * twenty mice, and a page that reported it as one item would tell IT they own
 * a fifth of what is in the cupboard.
 */

const unitsOf = (asset) => (Number.isFinite(asset?.quantity) ? asset.quantity : 1);

export function assetStats(assets = []) {
  const byCategory = new Map();
  const byStatus = new Map();
  const byLocation = new Map();

  let units = 0;
  let trackedUnits = 0;
  let unlabelled = 0;
  let faulty = 0;

  for (const asset of assets) {
    const count = unitsOf(asset);
    units += count;

    add(byCategory, asset.category || 'Uncategorised', count);
    add(byStatus, asset.status || 'In stock', count);
    if (asset.location) add(byLocation, asset.location, count);

    if (asset.trackingMode === TRACKED) {
      trackedUnits += count;
      // Only tracked things wear a sticker; a bag of cables was never going to.
      if (!String(asset.assetTag ?? '').trim()) unlabelled += 1;
    }

    if (asset.condition === 'Faulty') faulty += count;
  }

  return {
    rows: assets.length,
    units,
    trackedUnits,
    bulkUnits: units - trackedUnits,
    unlabelled,
    faulty,
    inStock: byStatus.get('In stock') ?? 0,
    byCategory: ranked(byCategory),
    byStatus: ranked(byStatus),
    byLocation: ranked(byLocation),
  };
}

function add(map, key, count) {
  map.set(key, (map.get(key) ?? 0) + count);
}

/** Biggest first, then alphabetically, so equal counts do not shuffle per render. */
function ranked(map) {
  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label));
}

/**
 * What arrived recently, for the "recent deliveries" strip. Grouped by
 * delivery rather than by item: thirty rows from one PO is one event.
 */
export function recentDeliveries(assets = [], limit = 5) {
  const byBatch = new Map();

  for (const asset of assets) {
    const key = asset.batchId || asset.batchTitle;
    if (!key) continue;

    const entry = byBatch.get(key) ?? {
      batchId: asset.batchId,
      title: asset.batchTitle || 'Delivery',
      supplier: asset.supplier || '',
      arrivedOn: asset.arrivedOn ?? null,
      units: 0,
      rows: 0,
    };

    entry.units += unitsOf(asset);
    entry.rows += 1;
    if ((asset.arrivedOn ?? 0) > (entry.arrivedOn ?? 0)) entry.arrivedOn = asset.arrivedOn;
    byBatch.set(key, entry);
  }

  return [...byBatch.values()]
    .sort((a, b) => (b.arrivedOn ?? 0) - (a.arrivedOn ?? 0))
    .slice(0, limit);
}
