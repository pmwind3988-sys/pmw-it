import { draftFromCodes, newDraft, setDraftField, draftIssues } from '../draft/draftAsset.js';
import { addDraft, replaceDraft } from '../draft/batch.js';
import { normaliseCode, indexByTag } from '../identity.js';
import { TRACKED, trackingModeFor } from '../assetKinds.js';
import { serialScore, partScore } from '../scan/classifyCode.js';
import { unitsOf, parseUnits, serialiseUnits } from '../units.js';

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
    // A counted line keeps its serials and labels on its ITEMS. One of those
    // read again is a thing already registered, not a new one.
    if (asset.trackingMode !== TRACKED) {
      for (const unit of unitsOf(asset)) {
        if (normaliseCode(unit.serialNumber) === wanted) return { asset, by: 'serial' };
        if (normaliseCode(unit.assetTag) === wanted) return { asset, by: 'tag' };
      }
    }
  }
  for (const asset of assets) {
    if (
      normaliseCode(asset.partNumber) === wanted
      || (asset.additionalCodes ?? []).some((code) => normaliseCode(code) === wanted)
      // A counted line's box barcode was filed under its first item when the
      // line was saved, so the next delivery of the same box finds it there.
      || (asset.trackingMode !== TRACKED
        && unitsOf(asset).some((unit) => normaliseCode(unit.partNumber) === wanted))
    ) {
      return { asset, by: 'part' };
    }
  }
  return null;
}

const union = (...lists) => [...new Set(lists.flat().map((code) => String(code ?? '').trim()).filter(Boolean))];

/**
 * A model draft that remembers the box barcode it was recognised by. On a
 * counted line the code goes into the row's OTHER CODES — a part number on a
 * counted line is moved onto one item when saved, and then the next box of
 * the same model would not be recognised from the line itself.
 */
function draftRemembering(asset, code, overrides = {}) {
  if (asset.trackingMode === TRACKED) {
    return draftOfModel(asset, { partNumber: code || asset.partNumber || '', ...overrides });
  }
  return draftOfModel(asset, {
    partNumber: '',
    additionalCodes: union(asset.additionalCodes ?? [], code ? [code] : []),
    quantity: 1,
    ...overrides,
  });
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
    // Carried so a save does not drop the barcodes the line is known by.
    additionalCodes: [...(asset.additionalCodes ?? [])],
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
    && !(draft.manualFields ?? []).includes('category');
}

const sameModel = (a, b) => ['category', 'manufacturer', 'model']
  .every((field) => String(a[field] ?? '').trim().toLowerCase() === String(b[field] ?? '').trim().toLowerCase());

/**
 * A model picked rather than scanned — from type-to-find or a quick key. A
 * bulk model already on the receipt counts up, exactly as scanning its box
 * again would; a tracked one is another machine, waiting for its serial.
 */
export function addModel(batch, asset) {
  if (asset.trackingMode !== TRACKED) {
    const existing = batch.drafts.find((draft) => draft.trackingMode !== TRACKED && sameModel(draft, asset));
    if (existing) {
      const next = setDraftField(existing, 'quantity', (existing.quantity ?? 1) + 1);
      return { batch: replaceDraft(batch, next), result: STOCK_RESULT.COUNTED, draft: next };
    }
    const draft = draftOfModel(asset, { quantity: 1, scanSource: 'Manual' });
    return { batch: addDraft(batch, draft), result: STOCK_RESULT.ADDED, draft };
  }
  const draft = draftOfModel(asset, { scanSource: 'Manual' });
  return { batch: addDraft(batch, draft), result: STOCK_RESULT.ADDED, draft };
}

/** A tracked line from a part number: one specific machine nobody has named. */
export function needsSerial(draft) {
  // A count line whose serial was looked for and not found says so outright;
  // asking again for what is not there would hold the count for ever.
  if (draft.noSerial) return false;
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
    // The same model already on the receipt — recognised by a different box
    // barcode, or linked a moment ago — is the same line: one more, and the
    // line now remembers this barcode too.
    const same = match.asset.trackingMode !== TRACKED
      && batch.drafts.find((draft) => draft.trackingMode !== TRACKED && !needsKind(draft) && sameModel(draft, match.asset));
    if (same) {
      const next = {
        ...setDraftField(same, 'quantity', (same.quantity ?? 1) + 1),
        additionalCodes: union(same.additionalCodes ?? [], [code]),
      };
      return { batch: replaceDraft(batch, next), result: STOCK_RESULT.COUNTED, draft: next };
    }
    const draft = draftRemembering(match.asset, code);
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
  if (!stopping) return null;
  // The review grid explains why at length; at a till the person needs the
  // instruction, not the reasoning.
  return stopping.field === 'model' ? 'Type its make and model.' : stopping.message;
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


/** The code an unknown line was scanned by — what it should be remembered as. */
function codeOf(draft) {
  return draft.partNumber || draft.serialNumber || (draft.additionalCodes ?? [])[0] || '';
}

/**
 * "Same as one we have": an unknown code is a model the register already
 * holds. The line becomes that model, and on a counted line the box barcode is
 * remembered against it — so every later box, in this delivery or the next,
 * is recognised on sight instead of asking again.
 */
export function linkToModel(batch, localId, asset) {
  const draft = batch.drafts.find((entry) => entry.localId === localId);
  if (!draft) return batch;
  const code = codeOf(draft);
  let next;
  if (asset.trackingMode === TRACKED) {
    // A tracked box's own code is usually its serial; a part-shaped one is
    // the model's, and the machine then waits for its serial.
    next = looksLikeSerial(code)
      ? draftOfModel(asset, { serialNumber: code, partNumber: asset.partNumber ?? '' })
      : draftOfModel(asset, { partNumber: code });
  } else {
    next = draftRemembering(asset, code, { quantity: draft.quantity ?? 1 });
  }
  return replaceDraft(batch, { ...next, localId: draft.localId, photoId: draft.photoId ?? null });
}

/**
 * A serial run onto a counted line: each serial is one item, in the next free
 * positions; `without` items were counted with no serial. The line's count
 * becomes at least what was scanned — scanning twelve serials off a line of
 * ten boxes means twelve.
 */
export function addSerialsTo(batch, localId, serials = [], without = 0) {
  const draft = batch.drafts.find((entry) => entry.localId === localId);
  if (!draft || draft.trackingMode === TRACKED) return batch;
  const units = parseUnits(draft.units);
  const taken = new Set(units.map((unit) => unit.index));
  let at = 0;
  const added = serials.map((serialNumber) => {
    while (taken.has(at)) at += 1;
    taken.add(at);
    return { index: at, serialNumber };
  });
  const recorded = units.length + added.length + Math.max(0, without);
  return replaceDraft(batch, {
    ...draft,
    units: serialiseUnits([...units, ...added]),
    quantity: Math.max(draft.quantity ?? 1, recorded),
  });
}

/** How many items on a counted line have a serial written on them. */
export function serialCount(draft) {
  return parseUnits(draft.units).filter((unit) => String(unit.serialNumber ?? '').trim()).length;
}

/**
 * What a box sweep found, onto a line. The category answers "what is this?"
 * on an unknown line; make and model fill only what is empty (a model the
 * register already named is not overwritten by a reading of the box); colour
 * and details join the line's details. Every field it writes is marked as
 * guessed, because it is.
 */
export function applySweep(batch, localId, found = {}) {
  let draft = batch.drafts.find((entry) => entry.localId === localId);
  if (!draft) return batch;
  if (found.category && needsKind(draft)) draft = setDraftField(draft, 'category', found.category);

  const guessed = new Set(draft.guessed ?? []);
  const fill = (field, value) => {
    if (!value || String(draft[field] ?? '').trim()) return;
    draft = { ...draft, [field]: value };
    guessed.add(field);
  };
  fill('manufacturer', found.make);
  fill('model', found.model);

  const extra = [found.colour, ...(found.details ?? [])].filter(Boolean).join(' · ');
  if (extra) {
    draft = { ...draft, specSummary: [draft.specSummary, extra].filter(Boolean).join(' · ') };
    guessed.add('specSummary');
  }
  return replaceDraft(batch, { ...draft, guessed: [...guessed] });
}
