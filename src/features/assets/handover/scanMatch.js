import { normaliseCode } from '../identity.js';
import { TRACKED } from '../assetKinds.js';
import { unitsOf } from '../units.js';

/**
 * What a scanned barcode points at: a whole item, or one item off a bulk row.
 *
 * The box in somebody's hand is one physical thing, and its barcode is on that
 * thing — so a code that matches a unit on a bulk row resolves to THAT unit,
 * not to the row as a whole. Scanning two tabs off a box of five therefore
 * lands on two different units, which is what lets the handover record two
 * items with two serials instead of one line of two.
 *
 * A tracked row IS one item, so it is never read as a unit — its serial belongs
 * to the row, and the row is what goes out. Unit matching is bulk rows only.
 *
 * Returns `{ asset, unit }` for a bulk item off the row, `{ asset, unit: null }`
 * for a whole row (tracked, or a bulk row matched on its own code), or `null`
 * when nothing in the register carries the code.
 */
export function findScanTarget(assets, rawCode) {
  const wanted = normaliseCode(rawCode);
  if (!wanted) return null;

  // A specific unit wins over the row it sits on: the more precise answer to
  // "which one is this" is the one worth having.
  for (const asset of assets) {
    if (asset.trackingMode === TRACKED) continue;
    for (const unit of unitsOf(asset)) {
      if (
        normaliseCode(unit.serialNumber) === wanted
        || normaliseCode(unit.assetTag) === wanted
        || normaliseCode(unit.partNumber) === wanted
      ) {
        return { asset, unit };
      }
    }
  }

  for (const asset of assets) {
    if (
      normaliseCode(asset.serialNumber) === wanted
      || normaliseCode(asset.assetTag) === wanted
      || normaliseCode(asset.partNumber) === wanted
      || (asset.additionalCodes ?? []).some((entry) => normaliseCode(entry) === wanted)
    ) {
      return { asset, unit: null };
    }
  }

  return null;
}
