import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME } from './deviceSchema.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';

/**
 * The only fields the register lets somebody retype. They are exactly the ones
 * the import DERIVED rather than read: everything else came out of the scan
 * file verbatim, and editing it here would put the register out of step with
 * the report that produced it.
 */
export const EDITABLE_FIELDS = ['owner', 'department', 'deviceType'];

const COLUMN_FOR = { owner: 'Owner', department: 'Department', deviceType: 'DeviceType' };

const asText = (value) => (value === null || value === undefined ? '' : String(value));

/**
 * Pure: what an edit changes, and what the row's manual list becomes.
 *
 * A field joins the manual list when it is edited, and leaves it when the
 * field is cleared. Blanking a value is how somebody hands it back to the scan
 * file — without that there would be no way to undo a correction, and the
 * field would stay frozen against every future import for good.
 */
export function planEdit(existing, edits) {
  const changes = [];
  const manual = new Set(existing.manualFields ?? []);

  for (const field of EDITABLE_FIELDS) {
    if (!(field in edits)) continue;

    const before = asText(existing[field]);
    const after = asText(edits[field]);
    if (before === after) continue;

    let changeType = 'Updated';
    if (!before) changeType = 'Added';
    else if (!after) changeType = 'Removed';

    changes.push({ fieldName: field, oldValue: before, newValue: after, changeType });

    if (after) manual.add(field);
    else manual.delete(field);
  }

  return { changes, manualFields: [...manual] };
}

function itemBody(edits, manualFields) {
  const body = { ManualFields: manualFields.join('\n') };

  for (const field of EDITABLE_FIELDS) {
    if (field in edits) body[COLUMN_FOR[field]] = asText(edits[field]);
  }

  // Owner Source stops claiming the value came from the scan once it did not.
  if ('owner' in edits) body.OwnerSource = manualFields.includes('owner') ? 'Manual' : 'Filename';

  return body;
}

async function logChanges(siteUrl, token, digest, computerName, changes, changedBy) {
  const changedOn = Date.now();

  for (const change of changes) {
    const response = await withRetry(() =>
      spFetch(siteUrl, `${listPath(CHANGE_LIST_NAME)}/items`, {
        token,
        digest,
        method: 'POST',
        accept: ITEM_ACCEPT,
        body: {
          Title: computerName,
          FieldName: change.fieldName,
          OldValue: change.oldValue,
          NewValue: change.newValue,
          ChangeType: change.changeType,
          ChangedOn: new Date(changedOn).toISOString(),
          ChangedOnMYT: formatMYT(changedOn, 'datetime12'),
          ChangedBy: changedBy ?? '',
        },
      }));

    if (!response.ok) {
      throw new Error(`Could not record the change (${response.status})`);
    }
  }
}

export async function updateDevice({
  siteUrl, token, existing, edits, changedBy,
}) {
  if (!existing?.id) throw new Error('That row has no id, so it cannot be updated');

  const { changes, manualFields } = planEdit(existing, edits);
  if (!changes.length) return { changes: [] };

  const digest = await getFormDigest(siteUrl, token);

  const response = await withRetry(() =>
    spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${existing.id})`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      body: itemBody(edits, manualFields),
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    }));

  if (!response.ok) {
    throw new Error(`Could not save the change (${response.status}): ${await response.text()}`);
  }

  await logChanges(siteUrl, token, digest, existing.computerName, changes, changedBy);
  return { changes };
}

export async function deleteDevice({
  siteUrl, token, device, changedBy,
}) {
  if (!device?.id) throw new Error('That row has no id, so it cannot be removed');

  const digest = await getFormDigest(siteUrl, token);

  const response = await withRetry(() =>
    spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${device.id})`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
    }));

  if (!response.ok) {
    throw new Error(`Could not remove the device (${response.status}): ${await response.text()}`);
  }

  // A machine leaving the register is never silent.
  await logChanges(siteUrl, token, digest, device.computerName, [{
    fieldName: 'device',
    oldValue: device.computerName,
    newValue: '',
    changeType: 'Removed',
  }], changedBy);

  return { removed: device.computerName };
}
