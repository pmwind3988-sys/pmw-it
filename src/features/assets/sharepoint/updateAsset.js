import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';
import { ASSET_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './assetSchema.js';
import { diffAsset } from './planSave.js';
import { assetKey, assetTitle } from '../identity.js';
import { TRACKED, BULK } from '../assetKinds.js';
import { diffUnits, withUnitsSplitOut } from '../units.js';
import { provisionAssets } from './provisionAssets.js';

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
  let next = { ...existing };

  for (const field of EDITABLE_FIELDS) {
    if (!(field in edits)) continue;
    next[field] = edits[field];

    if (asText(existing[field]) === asText(edits[field])) continue;
    if (asText(edits[field])) manual.add(field);
    else manual.delete(field);
  }

  // More than one of something is a line counted by quantity — the same rule
  // the review grid keeps (`setDraftField`), and it has to hold here too or a
  // delivery of ten monitors saved as one could be corrected to ten and save a
  // TRACKED row claiming a single serial for all ten. The flip must come
  // before the split below, which is what moves that serial onto item 1.
  if (Number(next.quantity) > 1 && next.trackingMode === TRACKED) {
    next.trackingMode = BULK;
    manual.add('trackingMode');
  }

  // A serial, a part number, a label, a condition and a status each describe
  // ONE physical item, so a bulk line does not keep them. On a row saved
  // before that rule they are moved onto item 1 in the same breath as being
  // cleared — separately would either lose them or leave one item's serial
  // standing in for twenty.
  next = withUnitsSplitOut(next);

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
  // Compared unconditionally, because the split above writes units nobody
  // edited — and a move from the row into item 1 has to show in the history as
  // the move it is, not as five fields vanishing.
  changes.push(...diffUnits(existing.units, next.units));

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

/**
 * SharePoint's way of saying "that column was never created here".
 *
 * The list is created once and then outlives every release, so a column added
 * to the app's schema afterwards — `Units` was — exists in the code and not in
 * the tenant. Every save of a row that touches it fails with this, and the
 * message it fails with ("The property 'Units' does not exist on type
 * 'SP.Data.IT_x0020_Asset_x0020_RegisterListItem'") tells the person holding
 * the phone nothing they can act on.
 */
export function isMissingColumn(status, body) {
  return status === 400 && /property '[^']+' does not exist on type/i.test(String(body ?? ''));
}

/**
 * The row write, and the one repair worth attempting automatically.
 *
 * Editing a row does not provision, on purpose: checking fifty columns before
 * every save would add a minute to each one. So the missing column is found
 * the only way that costs nothing on the ordinary path — by the save failing —
 * and provisioning runs then, once, before the same write is tried again.
 *
 * `repaired` is left false when nothing was wrong, so a caller can say "the
 * register was brought up to date" only when it actually was.
 */
async function writeRow(siteUrl, token, digest, id, body) {
  const send = () => withRetry(() => spFetch(siteUrl, `${listPath(ASSET_LIST_NAME)}/items(${id})`, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body,
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  }));

  const first = await send();
  if (first.ok) return { response: first, repaired: false };

  const failure = await first.text();
  if (!isMissingColumn(first.status, failure)) {
    throw new Error(`Could not save the change (${first.status}): ${failure}`);
  }

  await provisionAssets(siteUrl, token);

  const second = await send();
  if (!second.ok) {
    throw new Error(
      `Could not save the change (${second.status}): ${await second.text()}`,
    );
  }

  return { response: second, repaired: true };
}

export async function updateAsset({ siteUrl, token, existing, edits, changedBy }) {
  if (!existing?.id) throw new Error('That row has no id, so it cannot be updated');

  const { changes, record } = planEdit(existing, edits);
  // The unit records are compared separately from the change log, because a
  // photograph taken of item 3 changes the row and produces no log line — the
  // log deliberately ignores photos, and without this the save would decide
  // nothing had happened and quietly drop the picture.
  const unitsMoved = String(record.units ?? '') !== String(existing.units ?? '');
  if (!changes.length && !unitsMoved) return { changes: [] };

  const digest = await getFormDigest(siteUrl, token);

  const { repaired } = await writeRow(
    siteUrl, token, digest, existing.id, toListItem(record),
  );

  await logChanges(siteUrl, token, digest, record.assetKey, changes, changedBy);
  return { changes, record, repaired };
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
