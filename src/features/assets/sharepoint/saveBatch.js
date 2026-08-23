import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { runPool, withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';
import { provisionAssets } from './provisionAssets.js';
import { readAllAssets } from './readAssets.js';
import { planSave } from './planSave.js';
import { photoFolderUrl, uploadPhoto } from './uploadPhoto.js';
import {
  ASSET_LIST_NAME, BATCH_LIST_NAME, CHANGE_LIST_NAME, toListItem,
} from './assetSchema.js';
import { resolveDrafts, batchTitle } from '../draft/batch.js';

const itemPath = (listName) => `${listPath(listName)}/items`;

/**
 * A delivery leaving the phone for SharePoint.
 *
 * Progress is reported as `{ phase, done, total }` rather than a bare pair,
 * because the row writes are the short part. A first run spends over a minute
 * provisioning around fifty columns before anything moves, and a bar that sits
 * at "0 of 8" throughout is indistinguishable from a hang.
 */
export async function saveBatchToSharePoint({
  siteUrl, token, batch, photoFor, savedBy, onProgress,
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  // Provisioning runs to completion first and throws on failure: a half-created
  // list would fail every row with the same unhelpful message.
  report('provisioning');
  const digest = await provisionAssets(siteUrl, token, {
    onProgress: (done, total) => report('provisioning', done, total),
  });

  report('reading');
  const register = await readAllAssets(siteUrl, token);

  const savedOn = Date.now();
  const drafts = resolveDrafts(batch);
  const plan = planSave(drafts, register, { addedOn: savedOn, addedBy: savedBy ?? '' });

  const folder = await photoFolderUrl(siteUrl, token);

  // The PO scan first, because every row in the delivery points at it.
  report('photos');
  let poPhotoUrl = '';
  const poBlob = await photoFor?.(batch.purchase?.poPhotoId);
  if (poBlob) {
    poPhotoUrl = await uploadPhoto(siteUrl, token, digest, {
      folder, blob: poBlob, seed: batchTitle(batch), subfolder: 'po',
    }).catch(() => '');
  }

  const work = [
    ...plan.inserts.map((entry) => ({ ...entry, action: 'insert' })),
    ...plan.updates.map((entry) => ({ ...entry, action: 'update' })),
  ];

  /**
   * A photo that will not upload must not take its row down with it. The item
   * is the record; the photograph is an attachment to it, and losing the
   * serial number of a laptop because the camera produced an odd JPEG would be
   * the wrong trade. The failure is reported per row instead.
   */
  report('photos', 0, work.length);
  const photoUrls = new Map();
  const photoFailures = [];
  await runPool(work, async (entry) => {
    const blob = await photoFor?.(entry.body.photoId);
    if (!blob) return null;
    try {
      const url = await uploadPhoto(siteUrl, token, digest, {
        folder, blob, seed: entry.assetKey,
      });
      photoUrls.set(entry.localId, url);
    } catch (error) {
      photoFailures.push({ assetKey: entry.assetKey, error: error.message });
    }
    return null;
  }, { concurrency: 3, onProgress: (done, total) => report('photos', done, total) });

  report('writing', 0, work.length);
  const results = await runPool(work, async (entry) => {
    const body = toListItem({
      ...entry.body,
      photoUrl: photoUrls.get(entry.localId) ?? entry.body.photoUrl ?? '',
      poPhotoUrl: poPhotoUrl || (entry.body.poPhotoUrl ?? ''),
    });

    const response = entry.action === 'insert'
      ? await withRetry(() => spFetch(siteUrl, itemPath(ASSET_LIST_NAME), {
        token, digest, method: 'POST', body, accept: ITEM_ACCEPT,
      }))
      : await withRetry(() => spFetch(siteUrl, `${itemPath(ASSET_LIST_NAME)}(${entry.id})`, {
        token,
        digest,
        method: 'POST',
        body,
        accept: ITEM_ACCEPT,
        // A SharePoint update is a POST wearing these two headers.
        headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      }));

    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return entry.action;
  }, { concurrency: 4, onProgress: (done, total) => report('writing', done, total) });

  if (plan.changeRows.length) report('logging', 0, plan.changeRows.length);
  const changeResults = await runPool(plan.changeRows, async (row) => {
    const response = await withRetry(() => spFetch(siteUrl, itemPath(CHANGE_LIST_NAME), {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      body: {
        Title: row.assetKey,
        FieldName: row.fieldName,
        OldValue: row.oldValue,
        NewValue: row.newValue,
        ChangeType: row.changeType,
        ChangedOn: new Date(savedOn).toISOString(),
        ChangedOnMYT: formatMYT(savedOn, 'datetime12'),
        ChangedBy: savedBy ?? '',
      },
    }));
    if (!response.ok) throw new Error(String(response.status));
    return true;
  }, { concurrency: 4, onProgress: (done, total) => report('logging', done, total) });

  // The delivery's own row, written last: it records how many items it turned
  // out to contain, which is not known until the plan has run.
  report('delivery');
  const written = results.filter((result) => !result.error).length;
  await withRetry(() => spFetch(siteUrl, itemPath(BATCH_LIST_NAME), {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body: {
      Title: batchTitle(batch),
      Supplier: batch.purchase?.supplier ?? '',
      PoNumber: batch.purchase?.poNumber ?? '',
      ArrivedOn: instantOrNull(batch.purchase?.arrivedOn),
      ArrivedOnMYT: typeof batch.purchase?.arrivedOn === 'number'
        ? formatMYT(batch.purchase.arrivedOn, 'datetime12')
        : '',
      PoPhotoUrl: poPhotoUrl,
      ItemCount: written,
      Remarks: batch.purchase?.remarks ?? '',
      SavedOn: new Date(savedOn).toISOString(),
      SavedBy: savedBy ?? '',
    },
  }));

  return {
    results: results.map((result, index) => ({
      assetKey: work[index].assetKey,
      localId: work[index].localId,
      action: work[index].action,
      error: result.error ? result.error.message : null,
    })),
    blocked: plan.blocked,
    unchanged: plan.unchanged,
    changeCount: plan.changeRows.length,
    changeFailures: changeResults.filter((result) => result.error).length,
    photoFailures,
  };
}

function instantOrNull(value) {
  return typeof value === 'number' && Number.isFinite(value)
    ? new Date(value).toISOString()
    : null;
}

/**
 * Which rows still need saving after a partial failure.
 *
 * The batch keeps only what did not land, so pressing Save again does not
 * write the successful half a second time — the register would survive it
 * (the key upserts) but the change log would fill with phantom edits.
 */
export function remainingDrafts(batch, report) {
  const failed = new Set([
    ...report.results.filter((result) => result.error).map((result) => result.localId),
    ...report.blocked.map((entry) => entry.draft.localId),
  ]);

  return batch.drafts.filter((draft) => failed.has(draft.localId));
}
