import { assetKey, assetTitle, normaliseCode, indexByKey, indexByTag } from '../identity.js';
import { draftIssues, isBlocked } from '../draft/draftAsset.js';
import { TRACKED_FIELDS } from './assetSchema.js';
import { BULK } from '../assetKinds.js';
import { withUnitsSplitOut, appendUnit, mergeUnits, PER_UNIT_CODES } from '../units.js';

/**
 * What a save is going to do, decided before a single request is sent.
 *
 * Pure on purpose: every insert, update, quantity change and refusal is
 * testable without a token or a network, which is where the device import's
 * `planSync` earned its keep and where this one will too.
 */

/**
 * A field somebody corrected by hand outranks what a barcode says about it.
 * Without this the edit screen would be a trap: re-scanning the same box would
 * silently undo the correction.
 *
 * Only the named fields are held back, and the list itself is carried forward
 * so that updating anything else does not wipe it.
 */
export function applyManualOverrides(incoming, existing) {
  const manual = existing?.manualFields;
  if (!Array.isArray(manual) || !manual.length) return incoming;

  const merged = { ...incoming, manualFields: manual };
  for (const field of manual) {
    // A name left over from a field that no longer exists is skipped rather
    // than writing `undefined` over a real value.
    if (field in existing) merged[field] = existing[field];
  }
  return merged;
}

const asComparable = (value) => {
  if (value == null) return '';
  if (Array.isArray(value)) return value.join('\n');
  return String(value);
};

/** Field-level differences worth recording, in `TRACKED_FIELDS` order. */
export function diffAsset(existing, incoming) {
  const changes = [];

  for (const field of TRACKED_FIELDS) {
    const before = asComparable(existing?.[field]);
    const after = asComparable(incoming?.[field]);
    if (before === after) continue;

    changes.push({
      fieldName: field,
      oldValue: before,
      newValue: after,
      changeType: before === '' ? 'Added' : (after === '' ? 'Removed' : 'Updated'),
    });
  }

  return changes;
}

/**
 * Two rows in one batch that are the same thing.
 *
 * This is not a hypothetical: sweeping a shelf in many-items mode gives every
 * bag of the same mice its own row, and a second box of an already-scanned
 * model is a perfectly ordinary delivery. Bulk lines add up; tracked rows are
 * folded together, with the later row filling only what the earlier one lacks.
 */
export function coalesce(drafts) {
  const byKey = new Map();

  for (const draft of drafts) {
    const key = assetKey(draft);
    const existing = byKey.get(key);

    if (!existing) {
      // A bulk line's serial, part number, label and condition belong to the
      // box in front of the scanner, not to the line — so they come off the
      // row and become its first unit.
      byKey.set(key, withUnitsSplitOut({ ...draft, assetKey: key }, PER_UNIT_CODES));
      continue;
    }

    if (draft.trackingMode === BULK) {
      // The second of two identical tabs. Its codes take the NEXT position
      // rather than filling gaps in the first one's record: they describe a
      // different object, and merging them would invent an item that has one
      // tab's serial and the other's label.
      existing.units = appendUnit(existing.units, draft, existing.quantity ?? 0, PER_UNIT_CODES);
      existing.quantity = (existing.quantity ?? 0) + (draft.quantity ?? 0);
    }

    const shared = draft.trackingMode === BULK
      ? ['manufacturer', 'model', 'location', 'remarks', 'specSummary']
      : ['manufacturer', 'model', 'serialNumber', 'partNumber',
        'macAddress', 'assetTag', 'location', 'remarks', 'specSummary'];

    for (const field of shared) {
      if (!String(existing[field] ?? '').trim()) existing[field] = draft[field];
    }
    existing.photoId = existing.photoId ?? draft.photoId;
    existing.additionalCodes = [
      ...new Set([...(existing.additionalCodes ?? []), ...(draft.additionalCodes ?? [])]),
    ];
  }

  return [...byKey.values()];
}

/**
 * `drafts` are already resolved against their batch (§4.6), so the purchase
 * details are on them. `register` is every row currently in SharePoint.
 *
 * Returns `{ inserts, updates, blocked, changeRows, unchanged }`. A blocked row
 * never stops the others: a duplicate sticker label is one row's problem.
 */
export function planSave(drafts, register, { addedOn = Date.now(), addedBy = '' } = {}) {
  const byKey = indexByKey(register);
  const byTag = indexByTag(register);

  const inserts = [];
  const updates = [];
  const blocked = [];
  const changeRows = [];
  let unchanged = 0;

  // Labels claimed by rows earlier in this same batch, so two rows in one
  // delivery cannot both take PMW-0142.
  const claimedTags = new Map();

  for (const draft of coalesce(drafts)) {
    const issues = draftIssues(draft, { registerTags: byTag, batchTags: claimedTags });
    if (isBlocked(issues)) {
      blocked.push({ draft, issues });
      continue;
    }

    const tag = normaliseCode(draft.assetTag);
    if (tag) claimedTags.set(tag, draft.localId);

    const key = draft.assetKey ?? assetKey(draft);
    const existing = byKey.get(key);

    if (!existing) {
      inserts.push({
        localId: draft.localId,
        assetKey: key,
        body: { ...draft, assetKey: key, title: assetTitle(draft), addedOn, addedBy },
      });
      continue;
    }

    // A second bag of the same thing is more stock, not a correction to how
    // much stock there was. Tracked rows have no such arithmetic: there is one
    // of them by definition.
    const quantity = draft.trackingMode === BULK
      ? (existing.quantity ?? 0) + (draft.quantity ?? 0)
      : 1;

    // Diff and write the SAME record, so a field held back from the diff is
    // also held back from the body.
    // Unit records are carried forward explicitly, because the save writes
    // every column and a second bag of the same mice would otherwise erase
    // every serial anybody had recorded against the row.
    //
    // On a bulk line the box being scanned is a NEW physical item, so it takes
    // the next position rather than the first — the row already describes
    // however many arrived before it.
    const units = draft.trackingMode === BULK
      ? mergeUnits(existing.units, draft.units, existing.quantity ?? 0)
      : (draft.units ?? existing.units ?? '');

    const resolved = withUnitsSplitOut(applyManualOverrides(
      {
        ...draft,
        assetKey: key,
        quantity,
        status: draft.status ?? existing.status,
        units,
      },
      existing,
    ));

    const changes = diffAsset(existing, resolved);
    if (!changes.length) {
      unchanged += 1;
      continue;
    }

    updates.push({
      localId: draft.localId,
      assetKey: key,
      id: existing.id,
      body: { ...resolved, title: assetTitle(resolved) },
    });
    for (const change of changes) changeRows.push({ assetKey: key, ...change });
  }

  return { inserts, updates, blocked, changeRows, unchanged };
}
