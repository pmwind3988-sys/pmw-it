import { CATEGORIES, CONDITIONS, trackingModeFor, TRACKED } from '../assetKinds.js';
import { classifyCodes } from '../scan/classifyCode.js';
import { assetKey, hasStableIdentity, normaliseCode } from '../identity.js';

/**
 * One row of a batch, before it is anything in SharePoint.
 *
 * A draft is plain serialisable data — no Blobs, no class instances — because
 * it has to survive being written to IndexedDB and read back on a phone that
 * was closed and reopened. Photos live beside it as blobs keyed by `photoId`
 * for exactly that reason.
 */

/** `crypto.randomUUID` is everywhere this app runs; the fallback is for tests. */
export function newId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return `id-${Math.random().toString(36).slice(2)}-${Date.now().toString(36)}`;
}

export function newDraft(overrides = {}) {
  const category = overrides.category ?? 'Other';
  const trackingMode = overrides.trackingMode ?? trackingModeFor(category);

  return {
    localId: newId(),
    category,
    trackingMode,
    manufacturer: '',
    model: '',
    serialNumber: '',
    partNumber: '',
    macAddress: '',
    assetTag: '',
    // Tracked rows are one unit each by definition; the quantity box is only
    // editable on a bulk row, and `setDraftField` enforces it either way.
    quantity: 1,
    condition: CONDITIONS[0],
    location: '',
    remarks: '',
    specSummary: '',
    additionalCodes: [],
    photoId: null,
    // Per-row overrides of the batch's purchase details (§4.6). Absent means
    // "inherit", which is different from an empty string meaning "no supplier".
    supplier: undefined,
    poNumber: undefined,
    arrivedOn: undefined,
    scanSource: 'Camera',
    guessed: [],
    manualFields: [],
    ...overrides,
  };
}

/**
 * A draft built from the codes read off one box. Category is left at the
 * default deliberately: nothing in a barcode says what the thing is, and a
 * guessed category would silently decide whether the row is tracked or bulk.
 */
export function draftFromCodes(codes, overrides = {}) {
  const classified = classifyCodes(codes);

  return newDraft({
    serialNumber: classified.serialNumber,
    partNumber: classified.partNumber,
    macAddress: classified.macAddress,
    assetTag: classified.assetTag,
    additionalCodes: classified.additional,
    guessed: classified.guessed,
    ...overrides,
  });
}

const NUMERIC_FIELDS = new Set(['quantity']);

/**
 * Setting a field by hand, with the two bookkeeping consequences that must not
 * be forgotten anywhere: the field stops being a guess, and it starts
 * outranking a future re-scan (`manualFields`, the same contract the device
 * register uses).
 *
 * Changing the category re-derives the tracking mode — unless the mode was
 * itself set by hand, which is the override the spec promises.
 */
export function setDraftField(draft, field, value) {
  const next = { ...draft };

  if (NUMERIC_FIELDS.has(field)) {
    const parsed = Number(value);
    next[field] = Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : 1;
  } else {
    next[field] = value;
  }

  if (field === 'category' && !draft.manualFields.includes('trackingMode')) {
    next.trackingMode = trackingModeFor(value);
  }

  // A tracked row is one unit. Switching a bagful of cables to Tracked and
  // leaving the count at 20 would make twenty units share one serial.
  if (next.trackingMode === TRACKED) next.quantity = 1;

  next.guessed = draft.guessed.filter((name) => name !== field);
  next.manualFields = draft.manualFields.includes(field)
    ? draft.manualFields
    : [...draft.manualFields, field];

  return next;
}

/**
 * What is wrong with this row, in words a person can act on. Returned as a
 * list rather than a boolean so the review grid can show them all at once
 * instead of one per save attempt.
 *
 * `registerTags` maps a normalised label to the asset already wearing it.
 */
export function draftIssues(draft, { registerTags = new Map(), batchTags = new Map() } = {}) {
  const issues = [];

  if (!CATEGORIES.includes(draft.category)) {
    issues.push({ field: 'category', message: 'Pick a category.' });
  }

  if (!String(draft.model ?? '').trim() && !String(draft.serialNumber ?? '').trim()) {
    issues.push({
      field: 'model',
      message: 'Give it a model or a serial number — otherwise there is nothing to identify it by.',
    });
  }

  if (draft.trackingMode !== TRACKED && !(draft.quantity > 0)) {
    issues.push({ field: 'quantity', message: 'Quantity must be at least 1.' });
  }

  const key = assetKey(draft);
  const tag = normaliseCode(draft.assetTag);
  if (tag) {
    const owner = registerTags.get(tag);
    // A labelled machine being re-scanned finds its own row here. That is not
    // a clash — it is the ordinary case, and blocking it would make a labelled
    // asset the one kind that can never be updated.
    if (owner && owner.assetKey !== key) {
      issues.push({
        field: 'assetTag',
        message: `Label ${draft.assetTag} is already on "${owner.title ?? owner.assetKey}".`,
        blocking: true,
        conflictWith: owner.id ?? null,
      });
    }

    const twin = batchTags.get(tag);
    if (twin && twin !== draft.localId) {
      issues.push({
        field: 'assetTag',
        message: `Label ${draft.assetTag} is on another row in this batch.`,
        blocking: true,
      });
    }
  }

  if (!hasStableIdentity(key)) {
    // Not blocking: an unserialised spare is a real thing to own. But it will
    // never be recognised again, and silence about that is how duplicates breed.
    issues.push({
      field: 'serialNumber',
      message: 'No serial or label, so re-scanning this later will add a second row.',
    });
  }

  return issues;
}

export function isBlocked(issues) {
  return issues.some((issue) => issue.blocking);
}
