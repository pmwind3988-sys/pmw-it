import { TRACKED } from '../assetKinds.js';
import { available, out, HANDOVER_KIND, HANDOVER_STATUS } from './availability.js';

/**
 * Taking back a return that should not have happened.
 *
 * One press of Return (a line or "Return everything") stamps every row it
 * touched with the SAME `returnedOn`, so the most recent return for a person
 * is exactly the set of their rows carrying their latest `returnedOn`. Undoing
 * it puts those rows back to Out and puts the register copies back the way
 * `planHandover` writes them, from what the handover row still knows: who,
 * which kind, issued when, due when.
 *
 * Two things it cannot know, and says so rather than guessing:
 * - How much of a row THIS return took. `ReturnedQuantity` is a running total,
 *   so a row partly returned earlier and finished by this return reopens in
 *   full. Rare (the person page returns everything still out), and the
 *   confirm dialog says so.
 * - The item's condition before the return. A return writes the condition it
 *   came back in over the register's own, and the old value is not kept.
 */

/** The person's rows that the latest return touched. */
export function lastReturnOf(handovers, email) {
  const wanted = String(email ?? '').trim().toLowerCase();
  if (!wanted) return { returnedOn: null, rows: [] };

  const theirs = handovers.filter(
    (row) => String(row.personEmail ?? '').toLowerCase() === wanted
      && Number.isFinite(row.returnedOn)
      && (row.returnedQuantity ?? 0) > 0,
  );
  if (!theirs.length) return { returnedOn: null, rows: [] };

  const returnedOn = Math.max(...theirs.map((row) => row.returnedOn));
  return { returnedOn, rows: theirs.filter((row) => row.returnedOn === returnedOn) };
}

/**
 * `rows` are the handover rows to reopen (from `lastReturnOf`), read fresh.
 *
 * A line is refused when the thing has gone to somebody else since: a tracked
 * item now assigned to another person, or a bulk line without enough on the
 * shelf to put back. Undoing that would hand one laptop to two people.
 */
export function planUndoReturn(rows, register) {
  const assetById = new Map(register.map((asset) => [asset.id, asset]));

  const handoverUpdates = [];
  const blocked = [];
  // Several lines of one bulk box reopen against ONE register row, so their
  // quantities add up before anything is written (same reason as planReturn).
  const backByAsset = new Map();

  for (const row of rows) {
    const restoring = row.returnedQuantity ?? 0;
    if (restoring <= 0) continue;

    const asset = assetById.get(row.assetId);
    if (asset) {
      const already = backByAsset.get(asset.id);
      const pending = (already?.total ?? 0) + restoring;
      const ownEmail = String(row.personEmail ?? '').toLowerCase();
      const holder = String(asset.assignedToEmail ?? '').toLowerCase();

      if (asset.trackingMode === TRACKED && asset.assignedTo && holder !== ownEmail) {
        blocked.push({ row, reason: `${asset.assignedTo} has this now.` });
        continue;
      }
      if (asset.trackingMode !== TRACKED && pending > available(asset)) {
        blocked.push({ row, reason: `Only ${available(asset)} left on the shelf.` });
        continue;
      }

      backByAsset.set(asset.id, { asset, total: pending, row });
    }

    handoverUpdates.push({
      id: row.id,
      assetKey: row.assetKey,
      body: {
        returnedQuantity: 0,
        handoverStatus: HANDOVER_STATUS.OUT,
        returnedOn: null,
        returnedBy: '',
        returnCondition: null,
        // Cleared so the signature from a checklist link can land here later:
        // the server only writes a signature where there is none.
        returnSignature: '',
      },
    });
  }

  const assetUpdates = [...backByAsset.values()].map(({ asset, total, row }) => {
    const tracked = asset.trackingMode === TRACKED;
    const borrowed = row.kind === HANDOVER_KIND.BORROWED;
    return {
      id: asset.id,
      assetKey: asset.assetKey,
      body: {
        quantityOut: tracked ? 1 : out(asset) + total,
        // Exactly what planHandover writes, with the original dates.
        ...(tracked ? {
          status: borrowed ? 'Borrowed' : 'Assigned',
          handoverKind: row.kind ?? HANDOVER_KIND.ISSUED,
          assignedTo: row.personName ?? '',
          assignedToEmail: row.personEmail ?? '',
          assignedOn: row.issuedOn ?? null,
          dueOn: borrowed ? (row.dueOn ?? null) : null,
        } : {}),
      },
    };
  });

  return { handoverUpdates, assetUpdates, blocked };
}
