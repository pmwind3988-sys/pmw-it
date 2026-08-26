import { newId } from './draftAsset.js';

/**
 * A delivery, held on the phone until somebody reviews it.
 *
 * The purchase details — who it came from, which PO, when it arrived, the
 * photo of the paperwork — belong to the delivery rather than to each item in
 * it. Typing the supplier thirty times is how a register stops being filled
 * in, so it is typed once and copied down (§4.6).
 *
 * "Copied down" and not "looked up": each saved row carries its own supplier
 * and PO number, so a row read directly in SharePoint is complete on its own.
 */

export const BATCH_STATUS = { OPEN: 'open', REVIEWING: 'reviewing', SAVED: 'saved' };

function newPurchase() {
  return {
    supplier: '',
    poNumber: '',
    // The delivery order number off the note that came with the boxes. Blank
    // on most deliveries, and the whole point of a backfilled one.
    doNumber: '',
    arrivedOn: Date.now(),
    // A delivery entered long after it arrived, with its paperwork gone. See
    // `detailsPending.js` for why this is one switch and not thirty.
    detailsPending: false,
    poPhotoId: null,
    remarks: '',
  };
}

export function newBatch(overrides = {}) {
  return {
    id: newId(),
    createdAt: Date.now(),
    status: BATCH_STATUS.OPEN,
    drafts: [],
    ...overrides,
    // Merged rather than replaced, so a caller passing only `supplier` does not
    // silently drop the arrival date every row inherits.
    purchase: { ...newPurchase(), ...(overrides.purchase ?? {}) },
  };
}

/**
 * `undefined` on a draft means "inherit"; an empty string means "this row
 * genuinely has no supplier". Keeping those apart is what lets one line of a
 * delivery note come from somewhere else without the other rows following it.
 */
function inherit(rowValue, batchValue) {
  return rowValue === undefined ? batchValue : rowValue;
}

/** A draft with the delivery's details filled in where it has none of its own. */
export function resolveDraft(draft, batch) {
  const purchase = batch?.purchase ?? {};
  return {
    ...draft,
    supplier: inherit(draft.supplier, purchase.supplier ?? ''),
    poNumber: inherit(draft.poNumber, purchase.poNumber ?? ''),
    doNumber: inherit(draft.doNumber, purchase.doNumber ?? ''),
    detailsPending: inherit(draft.detailsPending, purchase.detailsPending ?? false),
    arrivedOn: inherit(draft.arrivedOn, purchase.arrivedOn ?? null),
    batchId: batch?.id ?? null,
    batchTitle: batchTitle(batch),
  };
}

export function resolveDrafts(batch) {
  return (batch?.drafts ?? []).map((draft) => resolveDraft(draft, batch));
}

/**
 * What the delivery is called. The PO number if there is one, because that is
 * what somebody looking for it will search; otherwise the supplier and the
 * date, which is at least sayable out loud.
 */
export function batchTitle(batch) {
  const po = String(batch?.purchase?.poNumber ?? '').trim();
  // Most PO numbers are typed with their own "PO" on the front, and "PO PO-4471"
  // reads as a mistake in the software rather than as a reference.
  if (po) return /^PO\b/i.test(po) ? po : `PO ${po}`;

  const supplier = String(batch?.purchase?.supplier ?? '').trim();
  const when = new Date(batch?.purchase?.arrivedOn ?? batch?.createdAt ?? Date.now());
  const day = Number.isNaN(when.getTime()) ? '' : when.toISOString().slice(0, 10);

  if (supplier && day) return `${supplier} — ${day}`;
  return supplier || day || 'Delivery';
}

export function addDraft(batch, draft) {
  return { ...batch, drafts: [...batch.drafts, draft] };
}

export function replaceDraft(batch, draft) {
  return {
    ...batch,
    drafts: batch.drafts.map((entry) => (entry.localId === draft.localId ? draft : entry)),
  };
}

export function removeDraft(batch, localId) {
  return { ...batch, drafts: batch.drafts.filter((draft) => draft.localId !== localId) };
}

/**
 * Two rows that turned out to be one box — the case a sweep in MANY mode
 * produces when a box carries both a serial and a part-number barcode.
 *
 * The kept row wins every field it has filled; the absorbed row contributes
 * only what the other one lacks, and its codes are never dropped.
 */
export function mergeDrafts(batch, keepId, absorbId) {
  const keep = batch.drafts.find((draft) => draft.localId === keepId);
  const absorb = batch.drafts.find((draft) => draft.localId === absorbId);
  if (!keep || !absorb || keepId === absorbId) return batch;

  const merged = { ...keep };
  for (const field of ['serialNumber', 'partNumber', 'macAddress', 'assetTag', 'model', 'manufacturer']) {
    if (!String(merged[field] ?? '').trim()) merged[field] = absorb[field];
  }

  const extra = [
    ...absorb.additionalCodes,
    // Anything the absorbed row carried that the kept row already had its own
    // answer for still belongs to the box; it goes to the codes list rather
    // than being thrown away.
    ...['serialNumber', 'partNumber', 'assetTag']
      .map((field) => absorb[field])
      .filter((value) => value && !Object.values(merged).includes(value)),
  ];

  merged.additionalCodes = [...new Set([...merged.additionalCodes, ...extra])];
  merged.photoId = merged.photoId ?? absorb.photoId;

  return {
    ...batch,
    drafts: batch.drafts
      .filter((draft) => draft.localId !== absorbId)
      .map((draft) => (draft.localId === keepId ? merged : draft)),
  };
}
