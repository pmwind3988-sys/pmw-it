import { spFetch, listPath, ITEM_ACCEPT, getFormDigest } from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { updateAsset, deleteAsset } from './updateAsset.js';
import { readAllHandovers } from './readHandovers.js';
import { HANDOVER_LIST_NAME, toUpdateItem } from './handoverSchema.js';
import { planCombine } from '../combine.js';

/**
 * What a handover on one of the folded rows has to say instead.
 *
 * `null` when it already says it — the keeper's own handovers usually do, and
 * a request that changes nothing is a request worth not sending.
 */
export function repointFor(move, kept) {
  const patch = {};

  if (kept.assetKey && move.was.assetKey !== kept.assetKey) patch.assetKey = kept.assetKey;
  if (kept.id && move.was.assetId !== kept.id) patch.assetId = kept.id;
  if (Number.isInteger(move.unitIndex) && move.was.unitIndex !== move.unitIndex) {
    patch.unitIndex = move.unitIndex;
  }

  /**
   * The serial of the row being folded away, written onto the handover.
   *
   * A tracked row handed out kept no serial of its own on the handover: the
   * row's TITLE said "Samsung S3 — 0XXXHNAL200474" and that was how everybody
   * read which monitor was on whose desk. Fold ten of those into one line and
   * the title becomes "Samsung S3", the same for all ten — so unless the
   * serial moves onto the handover here, the register stops being able to say
   * which one Iskandar has. It is written only when the handover has none:
   * a handover that named its own item knows better than the row does.
   */
  if (!move.was.serialNumber && move.row?.serialNumber) {
    patch.serialNumber = move.row.serialNumber;
  }

  // `itemTitle` is deliberately left alone. It is the record of WHAT WAS
  // HANDED OVER on the day, not a live pointer at a row — `assetKey` and
  // `assetId` are the pointer, and both have just been corrected. Rewriting
  // the title to the line's name would throw away the only wording that says
  // which of ten identical monitors went out.

  return Object.keys(patch).length ? patch : null;
}

/**
 * Several rows of the same thing written back as one line.
 *
 * The order is not an accident, and there are three steps now rather than two.
 *
 * The surviving row is written FIRST, carrying every serial number and label
 * the others held. Then the handovers are moved onto it — every person holding
 * one of these things goes on holding it, and now points at the item it has
 * become on the combined line. Only then are the other rows removed.
 *
 * A failure at any point therefore leaves MORE record rather than less: the
 * register holding everything twice, or a handover already pointing at the
 * line it is about to belong to. Both are visible and both can be finished by
 * hand. Removing the rows first would lose nine monitors between two requests,
 * and moving the handovers last would leave somebody holding an item whose row
 * had already gone.
 *
 * The handovers are re-read here rather than taken from the screen, for the
 * same reason a handover re-reads the register: somebody may have issued one of
 * these monitors from another phone while this page sat open, and that
 * handover has to move too.
 */
export async function combineAssets({
  siteUrl, token, rows, changedBy,
}) {
  const handovers = await readAllHandovers(siteUrl, token);
  const plan = planCombine(rows, handovers);

  const written = await updateAsset({
    siteUrl,
    token,
    existing: plan.keep,
    edits: plan.edits,
    changedBy,
  });

  const kept = written.record ?? plan.keep;
  const digest = await getFormDigest(siteUrl, token);

  let moved = 0;
  const failures = [];

  for (const move of plan.moves) {
    const patch = repointFor(move, kept);
    if (!patch || !move.was.id) continue;

    try {
      const response = await withRetry(() => spFetch(
        siteUrl,
        `${listPath(HANDOVER_LIST_NAME)}/items(${move.was.id})`,
        {
          token,
          digest,
          method: 'POST',
          body: toUpdateItem(patch),
          accept: ITEM_ACCEPT,
          headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
        },
      ));

      if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
      moved += 1;
    } catch (failure) {
      failures.push({ row: move.row, message: failure.message });
    }
  }

  /**
   * A handover left behind would point at a row about to be deleted, so the
   * rows it names are kept. Everything else still folds away — one stuck
   * record must not hold nine tidy ones hostage.
   */
  const stuck = new Set(failures.map((failure) => failure.row.id));

  const removeFailures = [];
  let removed = 0;
  for (const row of plan.remove) {
    if (stuck.has(row.id)) continue;

    try {
      await deleteAsset({ siteUrl, token, asset: row, changedBy });
      removed += 1;
    } catch (failure) {
      removeFailures.push({ row, message: failure.message });
    }
  }

  return {
    kept,
    quantity: plan.edits.quantity,
    moved,
    removed,
    failures: [...failures, ...removeFailures],
  };
}
