import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME } from './deviceSchema.js';
import { runPool, withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';
import { cleanLocation } from '../map/locations.js';
import { provisionDeviceColumns } from './provisionLists.js';

/**
 * The only fields the register lets somebody retype. They are exactly the ones
 * the import DERIVED rather than read: everything else came out of the scan
 * file verbatim, and editing it here would put the register out of step with
 * the report that produced it.
 */
export const EDITABLE_FIELDS = ['owner', 'department', 'deviceType', 'location'];

const COLUMN_FOR = { owner: 'Owner', department: 'Department', deviceType: 'DeviceType', location: 'Location' };

const asText = (value) => (value === null || value === undefined ? '' : String(value));
// Locations are stored as upper-case codes, however they were typed.
const editedText = (field, value) => (field === 'location' ? asText(cleanLocation(value)) : asText(value));

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
    const after = editedText(field, edits[field]);
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
    if (field in edits) body[COLUMN_FOR[field]] = editedText(field, edits[field]);
  }

  // Owner Source stops claiming the value came from the scan once it did not.
  if ('owner' in edits) body.OwnerSource = manualFields.includes('owner') ? 'Manual' : 'Filename';

  return body;
}

export async function logChanges(siteUrl, token, digest, device, changes, changedBy) {
  const changedOn = Date.now();

  for (const change of changes) {
    const response = await withRetry(() =>
      spFetch(siteUrl, `${listPath(CHANGE_LIST_NAME)}/items`, {
        token,
        digest,
        method: 'POST',
        accept: ITEM_ACCEPT,
        body: {
          Title: device.computerName ?? '',
          DeviceId: device.id,
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

  let digest = await getFormDigest(siteUrl, token);
  const merge = () => withRetry(() =>
    spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${existing.id})`, {
      token,
      digest,
      method: 'POST',
      accept: ITEM_ACCEPT,
      body: itemBody(edits, manualFields),
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    }));

  let response = await merge();

  // A list made before Location (or ManualFields) existed refuses the write by
  // naming the column. Only then are the columns added and the write tried
  // again -- checking first would cost every save the round trips that almost
  // never find anything.
  if (response.status === 400) {
    const reason = await response.text();
    if (!/does not exist on type/i.test(reason)) {
      throw new Error(`Could not save the change (400): ${reason}`);
    }
    digest = await provisionDeviceColumns(siteUrl, token);
    response = await merge();
  }

  if (!response.ok) {
    throw new Error(`Could not save the change (${response.status}): ${await response.text()}`);
  }

  await logChanges(siteUrl, token, digest, existing, changes, changedBy);
  return { changes };
}

/**
 * One machine off the register: the delete, then the row recording it. Shared
 * by the single Remove button and by removing a selection, so the two cannot
 * drift -- a removal logged by one route and not the other would leave the
 * change list lying about what the register holds.
 */
async function removeOne({
  siteUrl, token, digest, device, changedBy,
}) {
  if (device?.id === null || device?.id === undefined) {
    throw new Error('That row has no id, so it cannot be removed');
  }

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
  await logChanges(siteUrl, token, digest, device, [{
    fieldName: 'device',
    oldValue: device.computerName,
    newValue: '',
    changeType: 'Removed',
  }], changedBy);

  return device.computerName;
}

export async function deleteDevice({
  siteUrl, token, device, changedBy,
}) {
  if (device?.id === null || device?.id === undefined) {
    throw new Error('That row has no id, so it cannot be removed');
  }

  const digest = await getFormDigest(siteUrl, token);
  const removed = await removeOne({
    siteUrl, token, digest, device, changedBy,
  });
  return { removed };
}

/**
 * Several machines off the register in one go, four at a time and behind one
 * form digest.
 *
 * A machine that will not delete does NOT abandon the rest: somebody who ticked
 * twelve rows wants the eleven that can go to go, and to be told which one did
 * not. That is why this reports `{ removed, failures }` instead of throwing --
 * the only thing it throws for is a digest it could not get, which would fail
 * every row for the same reason.
 */
export async function deleteDevices({
  siteUrl, token, devices, changedBy, onProgress,
}) {
  if (!devices?.length) return { removed: [], failures: [] };

  const digest = await getFormDigest(siteUrl, token);

  const results = await runPool(
    devices,
    (device) => removeOne({
      siteUrl, token, digest, device, changedBy,
    }),
    { concurrency: 4, onProgress },
  );

  const removed = [];
  const failures = [];

  for (const result of results) {
    if (result.error) {
      failures.push({
        computerName: result.item?.computerName ?? '',
        error: result.error.message,
      });
    } else {
      removed.push(result.value);
    }
  }

  return { removed, failures };
}
