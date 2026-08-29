import { TRACKED, trackingModeFor } from '../assetKinds.js';
import { needsDetails } from '../detailsPending.js';
import { available, out as unitsOut } from '../handover/availability.js';
import { perItem, countPerItem, unitsOf as itemsOf } from '../units.js';

/**
 * The figures at the top of the register.
 *
 * Everything counts UNITS, not rows. A row reading "Logitech B100 × 20" is
 * twenty mice, and a page that reported it as one item would tell IT they own
 * a fifth of what is in the cupboard.
 */

const unitsOf = (asset) => (Number.isFinite(asset?.quantity) ? asset.quantity : 1);

/**
 * How many items on a bulk line are still waiting for a sticker.
 *
 * A bag of twenty cables was never going to wear twenty stickers, and counting
 * it would bury the number this card exists for. So a line is counted only
 * when one of two things is true:
 *
 *  - labelling has already STARTED on it — two tabs with one labelled is
 *    exactly the case somebody needs reminding about; or
 *  - it is a line of the KIND of thing that carries a label each. Monitors and
 *    laptops do, whatever `trackingMode` the row happens to be set to.
 *
 * That second rule is what makes the figure survive a combine. Ten monitors
 * scanned as ten tracked rows counted ten; folded into one line of ten they
 * counted ZERO, because nobody had started labelling the new line — so tidying
 * the register appeared to clear a job that nobody had done.
 */
function awaitingLabel(asset, count) {
  const tagged = itemsOf(asset).filter((unit) => unit.assetTag).length;
  if (tagged === 0 && trackingModeFor(asset?.category) !== TRACKED) return 0;
  return Math.max(0, count - tagged);
}

export function assetStats(assets = []) {
  const byCategory = new Map();
  const byStatus = new Map();
  const byLocation = new Map();

  let units = 0;
  let trackedUnits = 0;
  let unlabelled = 0;
  let faulty = 0;
  let out = 0;
  let free = 0;

  for (const asset of assets) {
    const count = unitsOf(asset);
    units += count;

    add(byCategory, asset.category || 'Uncategorised', count);
    if (asset.location) add(byLocation, asset.location, count);

    // Status is per item on a bulk line, so the tally comes from the items
    // rather than the row. Twenty cables with one retired is nineteen in stock
    // and one retired, not twenty of whichever the row happened to say.
    for (const entry of perItem(asset, 'status', 'In stock')) {
      add(byStatus, entry.value, entry.count);
    }

    if (asset.trackingMode === TRACKED) {
      trackedUnits += count;
      if (!String(asset.assetTag ?? '').trim()) unlabelled += 1;
    } else {
      unlabelled += awaitingLabel(asset, count);
    }

    faulty += countPerItem(asset, 'condition', 'Faulty');

    // Units, not rows: a box with three of twenty out contributes three, and
    // seventeen to what is still on the shelf.
    out += unitsOut(asset);
    free += available(asset);
  }

  return {
    rows: assets.length,
    units,
    trackedUnits,
    bulkUnits: units - trackedUnits,
    unlabelled,
    // ROWS, not units. A line of ten monitors is one delivery line to go back
    // and finish; counting the units would make the card read like a backlog
    // ten times the size of the job actually waiting.
    pending: assets.filter(needsDetails).length,
    faulty,
    out,
    available: free,
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
