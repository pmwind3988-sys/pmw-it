import { matchesQuery } from '../assetFilters.js';
import { available, isOpen, nameOfItem } from '../handover/availability.js';
import { TRACKED, BULK } from '../assetKinds.js';
import { unitsOf } from '../units.js';

/**
 * Type to find, for when there is no barcode worth scanning.
 *
 * Each mode searches what it could possibly mean: a delivery is more of a
 * MODEL the register knows, a handover is a row with something left, a return
 * is something somebody has — found by the thing OR by who has it.
 */

export const MIN_QUERY = 2;
const LIMIT = 6;

const has = (needle, ...values) => values.some((value) => String(value ?? '').toLowerCase().includes(needle));

function modelName(asset) {
  return [asset.manufacturer, asset.model].filter(Boolean).join(' ') || asset.title || 'Unnamed';
}

function searchModels(needle, assets) {
  const seen = new Set();
  const results = [];
  for (const asset of assets) {
    if (!String(asset.model ?? '').trim()) continue;
    const key = [asset.category, asset.manufacturer, asset.model].map((part) => String(part ?? '').toLowerCase()).join('|');
    if (seen.has(key)) continue;
    if (!has(needle, asset.model, asset.manufacturer, asset.category, asset.partNumber)) continue;
    seen.add(key);
    results.push({
      id: `model:${key}`,
      kind: 'model',
      name: modelName(asset),
      sub: `${asset.category || 'Other'} · ${asset.trackingMode === TRACKED ? 'one per serial' : 'counted in bulk'}`,
      asset,
    });
  }
  return results;
}

function searchStock(query, assets) {
  return assets
    .filter((asset) => available(asset) > 0 && matchesQuery(asset, query))
    .map((asset) => ({
      id: `asset:${asset.id}`,
      kind: 'asset',
      name: asset.title || modelName(asset),
      sub: asset.trackingMode === TRACKED
        ? [asset.serialNumber && `S/N ${asset.serialNumber}`, asset.assetTag].filter(Boolean).join(' · ') || 'In stock'
        : `${available(asset)} in stock`,
      asset,
    }));
}

function searchHolders(needle, assets, handovers) {
  const byKey = new Map(assets.map((asset) => [asset.assetKey, asset]));
  return handovers
    .filter((row) => isOpen(row))
    .filter((row) => {
      const asset = byKey.get(row.assetKey);
      return has(needle, row.itemTitle, row.serialNumber, row.personName, row.personEmail, asset?.title, asset?.assetTag, asset?.serialNumber);
    })
    .map((row) => {
      const asset = byKey.get(row.assetKey);
      const which = nameOfItem(row, asset ? unitsOf(asset) : []);
      return {
        id: `handover:${row.id}`,
        kind: 'handover',
        name: [row.itemTitle || asset?.title || 'Item', which].filter(Boolean).join(' · '),
        sub: `With ${row.personName || row.personEmail}`,
        handover: row,
      };
    });
}

export function tillSearch(mode, query, { assets = [], handovers = [] } = {}) {
  const trimmed = String(query ?? '').trim();
  if (trimmed.length < MIN_QUERY) return [];
  const needle = trimmed.toLowerCase();

  if (mode === 'in') return searchModels(needle, assets).slice(0, LIMIT);
  if (mode === 'out') return searchStock(trimmed, assets).slice(0, LIMIT);
  return searchHolders(needle, assets, handovers).slice(0, LIMIT);
}

/**
 * The bulk lines worth a key of their own on a desk: the ones handed out most
 * often, then the ones there are most of. Things with no barcode worth reading
 * are almost always bulk — a cable, a mouse — and those are what the keys save.
 */
export function quickKeys(assets = [], handovers = [], limit = 8) {
  const issued = new Map();
  for (const row of handovers) issued.set(row.assetKey, (issued.get(row.assetKey) ?? 0) + 1);

  return assets
    .filter((asset) => asset.trackingMode === BULK && String(asset.model ?? asset.title ?? '').trim())
    .map((asset) => ({ asset, issued: issued.get(asset.assetKey) ?? 0, left: available(asset) }))
    .sort((a, b) => (b.issued - a.issued) || ((b.asset.quantity ?? 0) - (a.asset.quantity ?? 0)))
    .slice(0, limit);
}
