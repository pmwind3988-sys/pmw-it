import { draftFromCodes, newDraft, setDraftField, draftIssues } from '../draft/draftAsset.js';
import { addDraft, replaceDraft } from '../draft/batch.js';
import { normaliseCode, indexByTag } from '../identity.js';
import { TRACKED, trackingModeFor } from '../assetKinds.js';
import { serialScore, partScore } from '../scan/classifyCode.js';

/**
 * Stock in at the till: what one scanned code does to the delivery on the
 * receipt.
 *
 * The receipt IS a batch (`draft/batch.js`), so everything downstream — the
 * IndexedDB copy, the review page, `planSave` — is unchanged. What the till
 * adds is the cashier's rule: the same thing scanned again counts up, and a
 * code the register already knows arrives already named.
 */

export const STOCK_RESULT = {
  EMPTY: 'empty',
  ADDED: 'added',
  COUNTED: 'counted',
  FILLED: 'filled',
  DUPLICATE: 'duplicate',
  KNOWN: 'known',
  NEW: 'new',
};

const PART_FIELDS = ['partNumber', 'additionalCodes'];

/** Every code a draft answers to, normalised, with the field it sits in. */
function codesOf(draft) {
  const out = [];
  for (const field of ['serialNumber', 'assetTag', 'macAddress', 'partNumber']) {
    const code = normaliseCode(draft[field]);
    if (code) out.push({ field, code });
  }
  for (const extra of draft.additionalCodes ?? []) {
    const code = normaliseCode(extra);
    if (code) out.push({ field: 'additionalCodes', code });
  }
  return out;
}

/**
 * The register row a code belongs to, and HOW it belongs. Serial and label
 * name one physical thing; a part number names a model, which is the
 * difference between "this laptop again" and "another one of these".
 * Whole rows only: an item inside a bulk line is not delivered twice.
 */
export function matchRegister(assets, raw) {
  const wanted = normaliseCode(raw);
  if (!wanted) return null;

  for (const asset of assets) {
    if (normaliseCode(asset.serialNumber) === wanted) return { asset, by: 'serial' };
    if (normaliseCode(asset.assetTag) === wanted) return { asset, by: 'tag' };
  }
  for (const asset of assets) {
    if (
      normaliseCode(asset.partNumber) === wanted
      || (asset.additionalCodes ?? []).some((code) => normaliseCode(code) === wanted)
    ) {
      return { asset, by: 'part' };
    }
  }
  return null;
}

/**
 * A draft of the same model as a register row. The descriptive fields are
 * carried across because a save writes every column: a draft with a blank
 * location would blank the row's.
 */
export function draftOfModel(asset, overrides = {}) {
  const category = asset.category || 'Other';
  return newDraft({
    category,
    trackingMode: asset.trackingMode || trackingModeFor(category),
    manufacturer: asset.manufacturer ?? '',
    model: asset.model ?? '',
    partNumber: asset.partNumber ?? '',
    specSummary: asset.specSummary ?? '',
    location: asset.location ?? '',
    scanSource: 'Camera',
    ...overrides,
  });
}

/**
 * A code nobody could name yet. Nothing in a barcode says what a thing is, so
 * the line asks — and checkout waits for the answer, because the category
 * decides whether the row is one item or a count.
 */
export function needsKind(draft) {
  return draft.category === 'Other'
    && !String(draft.model ?? '').trim()
    && !(draft.manualFields ?? []).includes('category');
}

/** A tracked line from a part number: one specific machine nobody has named. */
export function needsSerial(draft) {
  return draft.trackingMode === TRACKED
    && !String(draft.serialNumber ?? '').trim()
    && !String(draft.assetTag ?? '').trim();
}

function looksLikeSerial(raw) {
  return serialScore(raw) > partScore(raw);
}

/**
 * One code onto the receipt. Returns the new batch, what happened, and the
 * draft it happened to, so the page can say so in one line.
 */
export function scanIn(batch, raw, assets = [], format = '') {
  const code = normaliseCode(raw);
  if (!code) return { batch, result: STOCK_RESULT.EMPTY, draft: null };

  // 1. Already on this receipt.
  for (const draft of batch.drafts) {
    const hit = codesOf(draft).find((entry) => entry.code === code);
    if (!hit) continue;

    // A second bag of the same mice is one more; the same laptop's box read
    // twice is not a second laptop, and nor is an unnamed code read twice.
    const countable = draft.trackingMode !== TRACKED
      && !needsKind(draft)
      && PART_FIELDS.includes(hit.field);
    if (!countable) return { batch, result: STOCK_RESULT.DUPLICATE, draft };

    const next = setDraftField(draft, 'quantity', (draft.quantity ?? 1) + 1);
    return { batch: replaceDraft(batch, next), result: STOCK_RESULT.COUNTED, draft: next };
  }

  // 2. Known to the register.
  const match = matchRegister(assets, raw);
  if (match && match.by !== 'part') {
    // Already registered. Re-saving it from a delivery would rewrite the whole
    // row from what a barcode knows, so the till leaves it alone.
    return { batch, result: STOCK_RESULT.KNOWN, draft: null, asset: match.asset };
  }
  if (match) {
    const draft = match.asset.trackingMode === TRACKED
      ? draftOfModel(match.asset, { partNumber: code })
      : draftOfModel(match.asset, { quantity: 1 });
    return { batch: addDraft(batch, draft), result: STOCK_RESULT.ADDED, draft };
  }

  // 3. The serial of the machine whose part number was just read. A laptop box
  //    carries both, and reading them one after the other is one machine.
  const last = batch.drafts[batch.drafts.length - 1];
  if (last && needsSerial(last) && !needsKind(last) && looksLikeSerial(raw)) {
    const next = { ...last, serialNumber: code };
    return { batch: replaceDraft(batch, next), result: STOCK_RESULT.FILLED, draft: next };
  }

  // 4. Unknown.
  const draft = draftFromCodes([{ rawValue: String(raw).trim(), format }]);
  return { batch: addDraft(batch, draft), result: STOCK_RESULT.NEW, draft };
}

/** Naming an unknown line from its category chip. */
export function setKind(batch, localId, category) {
  const draft = batch.drafts.find((entry) => entry.localId === localId);
  if (!draft) return batch;
  return replaceDraft(batch, setDraftField(draft, 'category', category));
}

/** Any one field of a line, by hand. */
export function setLineField(batch, localId, field, value) {
  const draft = batch.drafts.find((entry) => entry.localId === localId);
  if (!draft) return batch;
  return replaceDraft(batch, setDraftField(draft, field, value));
}

const TAG_PATTERN = /^PMW-(\d+)$/;

/**
 * The next free sticker label: one above the highest PMW-NNNN in the register
 * or on this receipt, four digits. Counting from the highest rather than
 * filling gaps, because a gap is usually a sticker somebody has already
 * printed and not yet stuck on.
 */
export function nextTag(assets = [], drafts = []) {
  let highest = 0;
  const consider = (value) => {
    const found = TAG_PATTERN.exec(normaliseCode(value));
    if (found) highest = Math.max(highest, Number(found[1]));
  };
  for (const tag of indexByTag(assets).keys()) consider(tag);
  for (const draft of drafts) consider(draft.assetTag);
  return `PMW-${String(highest + 1).padStart(4, '0')}`;
}

/**
 * Something with nothing on it to scan. A tracked thing gets a new label so it
 * can be found next time; a bulk thing is identified by its model, as every
 * bulk line already is.
 */
export function noCodeDraft({ category, model = '', quantity = 1, tag = '' }) {
  const trackingMode = trackingModeFor(category);
  return newDraft({
    category,
    trackingMode,
    model: model.trim(),
    assetTag: trackingMode === TRACKED ? tag : '',
    quantity: trackingMode === TRACKED ? 1 : Math.max(1, Math.floor(Number(quantity) || 1)),
    scanSource: 'Manual',
    manualFields: ['category', 'model'],
  });
}

/**
 * Why a line cannot be saved yet, in words, or null. The till holds checkout
 * on these rather than letting `planSave` refuse them afterwards, because
 * the person who can answer is standing at the till now.
 */
export function holdOf(draft, { registerTags = new Map(), batchTags = new Map() } = {}) {
  if (needsKind(draft)) return 'Say what this is.';
  if (needsSerial(draft)) return 'Scan or type its serial number.';

  const issues = draftIssues(draft, { registerTags, batchTags });
  const stopping = issues.find((issue) => issue.blocking || issue.field === 'category' || issue.field === 'model');
  return stopping ? stopping.message : null;
}

/** Label → localId of the first line wearing it, the shape `draftIssues` wants. */
export function batchTagsOf(drafts) {
  const tags = new Map();
  for (const draft of drafts) {
    const tag = normaliseCode(draft.assetTag);
    if (tag && !tags.has(tag)) tags.set(tag, draft.localId);
  }
  return tags;
}

export function holdsFor(batch, assets = []) {
  const registerTags = indexByTag(assets);
  const batchTags = batchTagsOf(batch.drafts);
  const holds = new Map();
  for (const draft of batch.drafts) {
    const hold = holdOf(draft, { registerTags, batchTags });
    if (hold) holds.set(draft.localId, hold);
  }
  return holds;
}

export function itemCount(batch) {
  return batch.drafts.reduce((sum, draft) => sum + (draft.trackingMode === TRACKED ? 1 : (draft.quantity ?? 1)), 0);
}

