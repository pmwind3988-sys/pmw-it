import { CONDITIONS, trackingModeFor, TRACKED, BULK } from '../assetKinds.js';
import { classifyCodes } from '../scan/classifyCode.js';
import { assetKey, hasStableIdentity, normaliseCode } from '../identity.js';
import { needsDetails } from '../detailsPending.js';

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
    doNumber: undefined,
    arrivedOn: undefined,
    detailsPending: undefined,
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
  const manual = new Set(draft.manualFields);

  if (NUMERIC_FIELDS.has(field)) {
    const parsed = Number(value);
    next[field] = Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : 1;
  } else {
    next[field] = value;
  }

  // More than one of something is a line counted by quantity, whatever its
  // category usually says. Ten monitors delivered together are one line
  // reading ten, not ten rows — and the tracked-means-one-unit rule below is
  // kept rather than broken to get there: the ROW stops being tracked, and
  // each monitor's own serial goes to its own unit record (`units.js`), which
  // is exactly what the `Tab` category already does.
  //
  // Only ever in this direction. Bringing the count back down to one must NOT
  // flip it back, because lowering a quantity only HIDES units — pinning the
  // row to a single tracked unit would take the other nine serials with it.
  if (field === 'quantity' && next.quantity > 1 && next.trackingMode === TRACKED) {
    next.trackingMode = BULK;
    // Counted by hand, so a later category change cannot quietly undo it.
    manual.add('trackingMode');
  }

  if (field === 'category' && !manual.has('trackingMode')) {
    next.trackingMode = trackingModeFor(value);
  }

  // A tracked row is one unit. Switching a bagful of cables to Tracked and
  // leaving the count at 20 would make twenty units share one serial.
  if (next.trackingMode === TRACKED) next.quantity = 1;

  next.guessed = draft.guessed.filter((name) => name !== field);
  manual.add(field);
  next.manualFields = [...manual];

  return next;
}

/**
 * The two codes, the other way round.
 *
 * Which barcode on a box is the serial is a guess, and the guess is sometimes
 * wrong. Retyping both by hand off a label held in the other hand is the kind
 * of correction people skip, so the review grid offers the swap as one press —
 * and, being a correction, it marks both fields as set by hand, the same as
 * typing them would.
 */
export function swapSerialAndPart(draft) {
  return {
    ...draft,
    serialNumber: draft.partNumber ?? '',
    partNumber: draft.serialNumber ?? '',
    guessed: (draft.guessed ?? []).filter(
      (name) => name !== 'serialNumber' && name !== 'partNumber',
    ),
    manualFields: [...new Set([...(draft.manualFields ?? []), 'serialNumber', 'partNumber'])],
  };
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

  // Anything named counts. The built-in list is there to help somebody pick,
  // not to police the answer: a category can be added now (`categories.js`),
  // and a delivery of the first projector must not be refused for using it.
  if (!String(draft.category ?? '').trim()) {
    issues.push({ field: 'category', message: 'Pick a category.' });
  }

  // A serial identifies a tracked row and nothing else: a bulk line is keyed
  // on what the thing IS, and its serials belong to the individual items
  // inside it, so a bulk row with no model has nothing to be called at all.
  const named = String(draft.model ?? '').trim();
  const serialised = String(draft.serialNumber ?? '').trim();
  if (!named && (draft.trackingMode !== TRACKED || !serialised)) {
    issues.push({
      field: 'model',
      message: draft.trackingMode === TRACKED
        ? 'Give it a model or a serial number — otherwise there is nothing to identify it by.'
        : 'Give it a model. A line counted by quantity is identified by what it is, '
          + 'and its serials belong to the individual items on it.',
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

  // On a delivery whose paperwork is missing this warning is true and useless:
  // the serial is on a machine already sitting on somebody's desk, and saying
  // so on all thirty rows is how the whole flag gets ignored. A CLASH is still
  // a clash -- missing paperwork excuses a blank, never a collision.
  if (!hasStableIdentity(key) && !needsDetails(draft)) {
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
