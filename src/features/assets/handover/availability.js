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

/**
 * The open handovers on one item, gathered per PERSON.
 *
 * Somebody who took five cables on Monday and one more on Wednesday is two
 * rows in the handover list — correctly, because they were two events — and
 * one person on the shelf. Six identical lines under one name is a list nobody
 * reads; the question "who has this" wants one line per person and a total.
 *
 * Nothing is lost by the gathering: which item each line names is carried
 * through — by its own serial, or by looking up the unit it points at on a
 * bulk line — the soonest due date wins, and the person counts as overdue
 * if ANY of their lines is — one thing kept too long is kept too long, however
 * punctual the rest were.
 */
export function nameOfItem(handover, units = []) {
  if (handover?.serialNumber) return handover.serialNumber;

  const at = Number(handover?.unitIndex);
  if (!Number.isInteger(at) || at < 0) return '';

  const unit = units.find((entry) => entry.index === at);
  // The number is the honest last resort: an item nobody has written a serial
  // on is still a particular one of the line, and "item 3" is what the pager
  // calls it.
  return unit?.serialNumber || unit?.assetTag || `item ${at + 1}`;
}

export function groupHolders(open = [], units = []) {
  const byPerson = new Map();

  for (const row of open) {
    // Keyed on the email, like everything else about a person: two people
    // spell a display name two ways, and one person spells it three.
    const key = String(row.personEmail || row.personName || '').toLowerCase();

    const entry = byPerson.get(key) ?? {
      key,
      email: row.personEmail ?? '',
      name: row.personName || row.personEmail || 'Someone',
      units: 0,
      lines: 0,
      serials: [],
      kinds: [],
      overdue: false,
      dueOn: null,
      dueOnMYT: '',
    };

    entry.units += outstanding(row);
    entry.lines += 1;
    const named = nameOfItem(row, units);
    if (named) entry.serials.push(named);
    if (row.kind && !entry.kinds.includes(row.kind)) entry.kinds.push(row.kind);
    if (isOverdue(row)) entry.overdue = true;

    // The soonest deadline is the one worth showing: it is the next thing to
    // go wrong, and a later one says nothing while an earlier one is pending.
    if (Number.isFinite(row.dueOn) && (entry.dueOn === null || row.dueOn < entry.dueOn)) {
      entry.dueOn = row.dueOn;
      entry.dueOnMYT = row.dueOnMYT ?? '';
    }

    byPerson.set(key, entry);
  }

  return [...byPerson.values()].sort(
    (a, b) => Number(b.overdue) - Number(a.overdue)
      || b.units - a.units
      || a.name.localeCompare(b.name),
  );
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
