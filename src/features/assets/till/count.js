import { newDraft } from '../draft/draftAsset.js';
import { addDraft, replaceDraft, newBatch } from '../draft/batch.js';
import { normaliseCode } from '../identity.js';
import { TRACKED, trackingModeFor } from '../assetKinds.js';
import { matchRegister } from './stockIn.js';
import { serialiseUnits, mergeUnits } from '../units.js';

/**
 * The catch-up count: getting what IT already owns into the register, fast,
 * when most of it carries no barcode and no label.
 *
 * A count is a batch like any delivery (so it is kept on the device, survives
 * a reload and saves through `planSave`), but it is not a delivery: nothing
 * arrived, so there is no supplier and no arrival date to invent. Each line
 * says it came from a count in its remarks.
 *
 * Bulk things are counted by model — a drawer of fourteen mice is one line of
 * fourteen. A tracked thing is one line each, with its serial when it can say
 * one, and an explicit "no serial" plus a photo when it cannot.
 */

export const COUNT_REMARK = 'Added in a catch-up count';

export const COUNT_RESULT = {
  ADDED: 'added',
  COUNTED: 'counted',
  DUPLICATE: 'duplicate',
  KNOWN: 'known',
  INVALID: 'invalid',
};

export function newCount() {
  return newBatch({ purchase: { arrivedOn: null, remarks: 'Catch-up count' } });
}

const clean = (value) => String(value ?? '').trim();
const key = (value) => clean(value).toLowerCase();

/**
 * The models worth offering for a category: what this count has just used
 * first (the next thing on the shelf is usually the same as the last), then
 * what the register holds most of. Each make + model once.
 */
export function recentModels(category, assets = [], drafts = [], limit = 8) {
  const seen = new Set();
  const out = [];
  const offer = (manufacturer, model) => {
    if (!clean(model)) return;
    const id = `${key(manufacturer)}|${key(model)}`;
    if (seen.has(id)) return;
    seen.add(id);
    out.push({ manufacturer: clean(manufacturer), model: clean(model) });
  };

  for (const draft of [...drafts].reverse()) {
    if (draft.category === category) offer(draft.manufacturer, draft.model);
  }

  const tally = new Map();
  for (const asset of assets) {
    if (asset.category !== category || !clean(asset.model)) continue;
    const id = `${key(asset.manufacturer)}|${key(asset.model)}`;
    const entry = tally.get(id) ?? { manufacturer: asset.manufacturer, model: asset.model, n: 0 };
    entry.n += asset.trackingMode === TRACKED ? 1 : (asset.quantity ?? 1);
    tally.set(id, entry);
  }
  for (const entry of [...tally.values()].sort((a, b) => b.n - a.n)) offer(entry.manufacturer, entry.model);

  return out.slice(0, limit);
}

/** Places already in use, so a count in the same room is one tap. */
export function knownLocations(assets = [], drafts = []) {
  const all = [...drafts.map((draft) => draft.location), ...assets.map((asset) => asset.location)];
  return [...new Set(all.map(clean).filter(Boolean))].slice(0, 12);
}

/**
 * One counted entry onto the count. `entry`:
 *   { category, manufacturer, model, quantity, location, condition,
 *     serialNumber, noSerial, photoId }
 */
export function addCounted(batch, entry, assets = []) {
  const category = clean(entry.category);
  const model = clean(entry.model);
  if (!category || !model) return { batch, result: COUNT_RESULT.INVALID, draft: null };

  const trackingMode = trackingModeFor(category);
  const manufacturer = clean(entry.manufacturer);
  const location = clean(entry.location);

  if (trackingMode !== TRACKED) {
    // A serial run says how many by itself: one item per serial scanned, plus
    // the ones counted with no serial. Its serials become the line's items.
    const serials = (entry.serials ?? []).map(normaliseCode).filter(Boolean);
    const without = Math.max(0, Math.floor(Number(entry.without) || 0));
    const run = serials.length + without > 0;
    const quantity = run ? serials.length + without : Math.max(1, Math.floor(Number(entry.quantity) || 1));
    const runUnits = serialiseUnits(serials.map((serialNumber, index) => ({ index, serialNumber })));
    // Same model in the same place is the same line: counting a second drawer
    // of the same mice adds to it rather than starting another.
    const existing = batch.drafts.find((draft) => draft.trackingMode !== TRACKED
      && key(draft.category) === key(category)
      && key(draft.manufacturer) === key(manufacturer)
      && key(draft.model) === key(model)
      && key(draft.location) === key(location));
    if (existing) {
      const next = {
        ...existing,
        quantity: existing.quantity + quantity,
        // The new serials take the positions after everything already counted.
        units: runUnits ? mergeUnits(existing.units, runUnits, existing.quantity) : existing.units,
        specSummary: existing.specSummary || String(entry.specSummary ?? '').trim(),
      };
      return { batch: replaceDraft(batch, next), result: COUNT_RESULT.COUNTED, draft: next };
    }
    const draft = newDraft({
      category, trackingMode, manufacturer, model, quantity, location, units: runUnits,
      specSummary: String(entry.specSummary ?? '').trim(),
      scanSource: 'Manual', remarks: COUNT_REMARK, manualFields: ['category', 'model'],
    });
    return { batch: addDraft(batch, draft), result: COUNT_RESULT.ADDED, draft };
  }

  const serialNumber = normaliseCode(entry.serialNumber);
  if (serialNumber) {
    if (batch.drafts.some((draft) => normaliseCode(draft.serialNumber) === serialNumber)) {
      return { batch, result: COUNT_RESULT.DUPLICATE, draft: null };
    }
    const known = matchRegister(assets, serialNumber);
    if (known && known.by === 'serial') {
      return { batch, result: COUNT_RESULT.KNOWN, draft: null, asset: known.asset };
    }
  } else if (!entry.noSerial) {
    return { batch, result: COUNT_RESULT.INVALID, draft: null };
  }

  const draft = newDraft({
    category,
    trackingMode,
    manufacturer,
    model,
    serialNumber,
    location,
    condition: entry.condition || 'Good',
    photoId: entry.photoId ?? null,
    specSummary: String(entry.specSummary ?? '').trim(),
    scanSource: 'Manual',
    noSerial: !serialNumber,
    remarks: serialNumber ? COUNT_REMARK : `${COUNT_REMARK}. No serial could be found on it.`,
    manualFields: ['category', 'model'],
  });
  return { batch: addDraft(batch, draft), result: COUNT_RESULT.ADDED, draft };
}
