import { matchesQuery } from '../assetFilters.js';
import { available, isOpen, nameOfItem, titleOfHeld } from '../handover/availability.js';
import { TRACKED, BULK } from '../assetKinds.js';
import { unitsOf } from '../units.js';
import { normaliseCode } from '../identity.js';

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

/**
 * The items on a bulk line whose own serial or label holds what was typed,
 * minus the ones somebody already has. Typing "XW10" names one tab off a box
 * of twenty, so the answer is that tab and its FULL serial, not the box.
 */
function unitHits(query, asset, handovers) {
  if (asset.trackingMode === TRACKED) return [];
  const code = normaliseCode(query);
  if (!code) return [];
  const held = new Set(handovers
    .filter((row) => row.assetKey === asset.assetKey && isOpen(row) && row.unitIndex != null && row.unitIndex !== '')
    .map((row) => Number(row.unitIndex)));
  return unitsOf(asset)
    .filter((unit) => !held.has(unit.index))
    .filter((unit) => [unit.serialNumber, unit.assetTag].some((value) => normaliseCode(value).includes(code)))
    .map((unit) => ({
      id: `unit:${asset.id}:${unit.index}`,
      kind: 'unit',
      name: `${asset.title || modelName(asset)} · item ${unit.index + 1}`,
      sub: [unit.serialNumber && `S/N ${unit.serialNumber}`, unit.assetTag].filter(Boolean).join(' · '),
      asset,
      unit,
    }));
}

function searchStock(query, assets, handovers) {
  const results = [];
  for (const asset of assets) {
    if (available(asset) <= 0) continue;
    const units = unitHits(query, asset, handovers);
    results.push(...units);
    // The line itself only when it matches on its OWN words (name, model,
    // part number) — matched through an item's serial, the item is the answer.
    // An item somebody already has does not bring its line back in either.
    if (matchesQuery({ ...asset, units: '[]' }, query)) results.push(lineResult(asset));
  }
  return results;
}

function lineResult(asset) {
  return {
    id: `asset:${asset.id}`,
    kind: 'asset',
    name: asset.title || modelName(asset),
    sub: asset.trackingMode === TRACKED
      ? [asset.serialNumber && `S/N ${asset.serialNumber}`, asset.assetTag].filter(Boolean).join(' · ') || 'In stock'
      : `${available(asset)} in stock`,
    asset,
  };
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
        name: [titleOfHeld(row, asset) || 'Item', which].filter(Boolean).join(' · '),
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
  if (mode === 'out') return searchStock(trimmed, assets, handovers).slice(0, LIMIT);
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

/**
 * Hand out's overview: how many of each KIND of thing are on the shelf, then
 * which products make that up. "Keyboards · 7" is the question somebody at the
 * desk is asking; which of two keyboard lines the seven sit on is the second
 * question, and the handover still has to name one, so each category carries
 * its products. Products group by name, so five tracked laptops of one model
 * are one entry of 5 (`rows` holds them) and a bulk line is an entry of its own.
 * Categories with nothing left still show, at 0, so "we have none" is visible.
 */
export function stockByCategory(assets = []) {
  const categories = new Map();
  for (const asset of assets) {
    const category = String(asset.category ?? '').trim() || 'Other';
    const name = String(asset.title || asset.model || '').trim();
    if (!name) continue;
    const left = available(asset);
    const group = categories.get(category) ?? { category, left: 0, products: new Map() };
    group.left += left;
    const product = group.products.get(name.toLowerCase()) ?? { name, left: 0, rows: [] };
    product.left += left;
    product.rows.push(asset);
    group.products.set(name.toLowerCase(), product);
    categories.set(category, group);
  }
  return [...categories.values()]
    .map((group) => ({
      ...group,
      products: [...group.products.values()]
        .sort((a, b) => (b.left - a.left) || a.name.localeCompare(b.name)),
    }))
    .sort((a, b) => (b.left > 0) - (a.left > 0) || a.category.localeCompare(b.category));
}
