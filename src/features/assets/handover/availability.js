import { TRACKED } from '../assetKinds.js';

/**
 * How much of a thing is on the shelf.
 *
 * The register's `Quantity` is what the company OWNS and never moves when
 * something is handed out; `QuantityOut` counts what is with people. Available
 * is the difference (§4.1).
 *
 * Keeping it that way means a return only ever moves the derived figure, so a
 * handover somebody forgot to record cannot silently change how many cables the
 * company believes it bought.
 */

export const HANDOVER_STATUS = {
  OUT: 'Out',
  PARTLY: 'Partly returned',
  RETURNED: 'Returned',
};

export const HANDOVER_KIND = { ISSUED: 'Issued', BORROWED: 'Borrowed' };

const count = (value) => (Number.isFinite(value) ? value : 0);

export function owned(asset) {
  if (asset?.trackingMode === TRACKED) return 1;
  return Number.isFinite(asset?.quantity) ? asset.quantity : 1;
}

export function out(asset) {
  return count(asset?.quantityOut);
}

/** Never negative: a bad figure should read as "none left", not as a credit. */
export function available(asset) {
  return Math.max(0, owned(asset) - out(asset));
}

export function isOut(asset) {
  return available(asset) < owned(asset);
}

/** What is still with the person on one handover row. */
export function outstanding(handover) {
  return Math.max(0, count(handover?.quantity) - count(handover?.returnedQuantity));
}

export function isOpen(handover) {
  return outstanding(handover) > 0;
}

/**
 * Overdue means borrowed, still out, and past its date. An Issued item has no
 * date and can never be overdue — that is the whole difference between the two
 * kinds (§4.3).
 */
export function isOverdue(handover, now = Date.now()) {
  return isOpen(handover)
    && handover?.kind === HANDOVER_KIND.BORROWED
    && Number.isFinite(handover?.dueOn)
    && handover.dueOn < now;
}

export function statusFor(handover) {
  const returned = count(handover?.returnedQuantity);
  if (returned <= 0) return HANDOVER_STATUS.OUT;
  if (returned >= count(handover?.quantity)) return HANDOVER_STATUS.RETURNED;
  return HANDOVER_STATUS.PARTLY;
}

/** The open handovers for one asset, which is who currently holds it. */
export function holdersOf(handovers, assetKey) {
  return handovers.filter((row) => row.assetKey === assetKey && isOpen(row));
}

/** Everything one person currently holds, keyed on email rather than on name. */
export function heldBy(handovers, email) {
  const wanted = String(email ?? '').trim().toLowerCase();
  if (!wanted) return [];
  return handovers.filter(
    (row) => String(row.personEmail ?? '').toLowerCase() === wanted && isOpen(row),
  );
}

/**
 * Everyone with something out, for the people list. Counted in units so a
 * person holding three cables and a laptop reads as four things, not two rows.
 */
export function peopleWithItems(handovers, now = Date.now()) {
  const byEmail = new Map();

  for (const row of handovers) {
    if (!isOpen(row)) continue;
    const email = String(row.personEmail ?? '').toLowerCase();
    if (!email) continue;

    const entry = byEmail.get(email) ?? {
      email,
      name: row.personName || row.personEmail,
      units: 0,
      lines: 0,
      overdue: 0,
    };

    entry.units += outstanding(row);
    entry.lines += 1;
    if (isOverdue(row, now)) entry.overdue += 1;
    byEmail.set(email, entry);
  }

  return [...byEmail.values()].sort(
    (a, b) => b.overdue - a.overdue || b.units - a.units || a.name.localeCompare(b.name),
  );
}
