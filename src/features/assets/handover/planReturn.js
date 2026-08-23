import { TRACKED } from '../assetKinds.js';
import { outstanding, out, statusFor, HANDOVER_STATUS } from './availability.js';

/**
 * Taking things back.
 *
 * A return moves two rows: the handover row records how much came back and in
 * what condition, and the register row's `QuantityOut` comes down. For a
 * tracked item the copied assignment fields are cleared as well, or the
 * register would go on claiming somebody has a laptop that is back on the
 * shelf.
 */

/**
 * `returns` is `[{ handoverId, quantity, condition }]` where `handoverId` is
 * the handover row's list id.
 *
 * Returning MORE than is out is refused rather than clamped: the difference
 * between "two came back" and "three came back" is a real disagreement about
 * what happened, and quietly rounding it hides a miscount.
 */
export function planReturn(returns, handovers, register, {
  returnedOn = Date.now(), returnedBy = '',
} = {}) {
  const handoverById = new Map(handovers.map((row) => [row.id, row]));
  const assetById = new Map(register.map((asset) => [asset.id, asset]));

  const handoverUpdates = [];
  const blocked = [];
  // Two lines of the same box coming back in one action have to accumulate
  // against one register row, or the second write overwrites the first's
  // arithmetic and half the return is lost.
  const outByAsset = new Map();

  for (const entry of returns) {
    const handover = handoverById.get(entry.handoverId);

    if (!handover) {
      blocked.push({ entry, reason: 'That handover is no longer in the list.' });
      continue;
    }

    const stillOut = outstanding(handover);
    const wanted = Number.isFinite(entry.quantity) ? entry.quantity : stillOut;

    if (wanted <= 0) {
      blocked.push({ entry, reason: 'Nothing to return on that line.' });
      continue;
    }

    if (wanted > stillOut) {
      blocked.push({
        entry,
        reason: `Only ${stillOut} still out on that handover.`,
      });
      continue;
    }

    const returnedQuantity = (handover.returnedQuantity ?? 0) + wanted;
    const next = { ...handover, returnedQuantity };

    handoverUpdates.push({
      id: handover.id,
      assetKey: handover.assetKey,
      body: {
        returnedQuantity,
        handoverStatus: statusFor(next),
        returnedOn,
        returnedBy,
        returnCondition: entry.condition ?? '',
      },
    });

    const asset = assetById.get(handover.assetId);
    if (!asset) continue;

    const already = outByAsset.get(asset.id);
    const base = already ? already.remaining : out(asset);
    const remaining = Math.max(0, base - wanted);

    outByAsset.set(asset.id, {
      asset,
      remaining,
      // The condition of the last item recorded on this row wins. Returning a
      // good cable and a broken one together is genuinely ambiguous, and the
      // per-line return is how somebody says which is which.
      condition: entry.condition || already?.condition || '',
      fullyBack: remaining === 0,
    });
  }

  const assetUpdates = [...outByAsset.values()].map(({ asset, remaining, condition, fullyBack }) => ({
    id: asset.id,
    assetKey: asset.assetKey,
    body: {
      quantityOut: remaining,
      ...(condition ? { condition } : {}),
      // Cleared only when nothing of it is out any more. A tracked item is
      // always all-or-nothing; a bulk row with two of five still out must keep
      // reading as partly out.
      ...(fullyBack && asset.trackingMode === TRACKED ? {
        status: 'In stock',
        handoverKind: '',
        assignedTo: '',
        assignedToEmail: '',
        assignedOn: null,
        dueOn: null,
      } : {}),
      ...(fullyBack && asset.trackingMode !== TRACKED ? { status: 'In stock' } : {}),
    },
  }));

  return { handoverUpdates, assetUpdates, blocked };
}

/** Everything one person still holds, as the return list for an offboarding. */
export function returnEverything(handovers, condition = '') {
  return handovers
    .filter((row) => row.handoverStatus !== HANDOVER_STATUS.RETURNED && outstanding(row) > 0)
    .map((row) => ({
      handoverId: row.id,
      quantity: outstanding(row),
      condition,
    }));
}
