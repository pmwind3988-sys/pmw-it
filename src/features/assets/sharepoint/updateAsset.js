import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';
import { ASSET_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './assetSchema.js';
import { diffAsset } from './planSave.js';
import { assetKey, assetTitle } from '../identity.js';
import { diffUnits } from '../units.js';

/**
 * Editing and removing one asset.
 *
 * Unlike the device register — where the scan file is the source of truth and
 * only the derived fields may be retyped — everything here is editable. There
 * is no file to disagree with: a barcode said what it said, and a person
 * holding the thing knows better.
 */
export const EDITABLE_FIELDS = [
  'category', 'trackingMode', 'manufacturer', 'model', 'serialNumber', 'partNumber',
  'macAddress', 'assetTag', 'quantity', 'condition', 'status', 'location',
  'remarks', 'specSummary', 'supplier', 'poNumber', 'arrivedOn', 'units',
];

const asText = (value) => (value === null || value === undefined ? '' : String(value));

/**
 * Pure: what an edit changes, and what the row's manual list becomes.
 *
 * A field joins the manual list when it is edited and leaves it when cleared.
 * Blanking a value is how somebody hands it back to the scanner — without that
 * there would be no way to undo a correction, and the field would stay frozen
 * against every future scan for good.
 */
export function planEdit(existing, edits) {
  const changes = [];
  const manual = new Set(existing.manualFields ?? []);
  const next = { ...existing };

  for (const field of EDITABLE_FIELDS) {
    if (!(field in edits)) continue;
    next[field] = edits[field];

    if (asText(existing[field]) === asText(edits[field])) continue;
    if (asText(edits[field])) manual.add(field);
    else manual.delete(field);
  }

  // `units` is a JSON blob, so it is held out of the manual list and out of
  // the ordinary diff. Both would be nonsense: nothing re-scans a unit record,
  // and a change log line reading `[{"index":1,...}]` is the same as no
  // history at all. It is logged below, one line per unit and field.
  manual.delete('units');
  next.manualFields = [...manual];
  // Both derived, both re-derived: correcting a serial number changes which
  // physical thing this row claims to be, and the key has to follow it or the
  // next scan of that machine makes a second row.
  next.assetKey = assetKey(next);
  next.title = assetTitle(next);

  changes.push(...diffAsset(existing, next));
  if ('units' in edits) changes.push(...diffUnits(existing.units, next.units));

  return { changes, record: next };
}

async function logChanges(siteUrl, token, digest, key, changes, changedBy) {
  const changedOn = Date.now();

  for (const change of changes) {
    const response = await withRetry(() => spFetch(siteUrl, `${listPath(CHANGE_LIST_NAME)}/items`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      body: {
        Title: key,
        FieldName: change.fieldName,
        OldValue: change.oldValue,
        NewValue: change.newValue,
        ChangeType: change.changeType,
        ChangedOn: new Date(changedOn).toISOString(),
        ChangedOnMYT: formatMYT(changedOn, 'datetime12'),
        ChangedBy: changedBy ?? '',
      },
    }));

    if (!response.ok) throw new Error(`Could not record the change (${response.status})`);
  }
}

export async function updateAsset({ siteUrl, token, existing, edits, changedBy }) {
  if (!existing?.id) throw new Error('That row has no id, so it cannot be updated');

  const { changes, record } = planEdit(existing, edits);
  if (!changes.length) return { changes: [] };

  const digest = await getFormDigest(siteUrl, token);

  const response = await withRetry(() =>
    spFetch(siteUrl, `${listPath(ASSET_LIST_NAME)}/items(${existing.id})`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      body: toListItem(record),
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    }));

  if (!response.ok) {
    throw new Error(`Could not save the change (${response.status}): ${await response.text()}`);
  }

  await logChanges(siteUrl, token, digest, record.assetKey, changes, changedBy);
  return { changes, record };
}

export async function deleteAsset({ siteUrl, token, asset, changedBy }) {
  if (!asset?.id) throw new Error('That row has no id, so it cannot be removed');

  const digest = await getFormDigest(siteUrl, token);

  const response = await withRetry(() =>
    spFetch(siteUrl, `${listPath(ASSET_LIST_NAME)}/items(${asset.id})`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
    }));

  if (!response.ok) {
    throw new Error(`Could not remove the item (${response.status}): ${await response.text()}`);
  }

  // Something leaving the register is never silent.
  await logChanges(siteUrl, token, digest, asset.assetKey, [{
    fieldName: 'asset',
    oldValue: asset.title ?? asset.assetKey,
    newValue: '',
    changeType: 'Removed',
  }], changedBy);

  return { removed: asset.title ?? asset.assetKey };
}
