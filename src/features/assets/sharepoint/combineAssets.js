import { updateAsset, deleteAsset } from './updateAsset.js';
import { planCombine } from '../combine.js';

/**
 * Several rows of the same thing written back as one line.
 *
 * The order is not an accident. The surviving row is written FIRST, carrying
 * every serial number and label the others held; only then are the others
 * removed. A failure halfway through therefore leaves the register holding
 * everything twice, which somebody can see and finish — the other order would
 * lose nine monitors between two requests.
 *
 * Both halves go through the ordinary edit and delete, so the change log reads
 * the way it always does: the line grew from one item to ten, and nine rows
 * were removed, each with its own line and its own name against it.
 */
export async function combineAssets({
  siteUrl, token, rows, changedBy,
}) {
  const plan = planCombine(rows);

  const written = await updateAsset({
    siteUrl,
    token,
    existing: plan.keep,
    edits: plan.edits,
    changedBy,
  });

  const failures = [];
  for (const row of plan.remove) {
    try {
      await deleteAsset({ siteUrl, token, asset: row, changedBy });
    } catch (failure) {
      failures.push({ row, message: failure.message });
    }
  }

  return {
    kept: written.record ?? plan.keep,
    quantity: plan.edits.quantity,
    removed: plan.remove.length - failures.length,
    failures,
  };
}
