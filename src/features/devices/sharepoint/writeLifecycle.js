import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { DEVICE_LIST_NAME } from './deviceSchema.js';
import { ASSIGNMENT_LIST_NAME, toAssignmentItem, closeBody } from './assignmentSchema.js';
import { readDevice } from './readDevices.js';
import { readAssignments } from './readAssignments.js';
import { logChanges } from './updateDevice.js';
import { planLifecycle } from '../lifecycle/planLifecycle.js';
import { provisionLists } from './provisionLists.js';

const MERGE = { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };

const TEXT_COLUMN = {
  owner: 'Owner', location: 'Location', department: 'Department', ownerSource: 'OwnerSource', status: 'Status',
};

// The columns these writes name only exist once an import has provisioned them,
// and the map -> machine -> action path can be the very first thing anyone does.
// One run per page load; a failed run is forgotten so the next press retries.
let provisioning = null;
// Resolves to the digest provisioning obtained only for the call that ran it;
// later calls get null and fetch their own, because a digest expires (~30 min)
// and a page left open must not keep sending a dead one.
async function ensureProvisioned(siteUrl, token) {
  if (provisioning) {
    await provisioning;
    return null;
  }
  const run = provisionLists(siteUrl, token);
  provisioning = run.catch((failure) => {
    provisioning = null;
    throw failure;
  });
  return (await provisioning, await run);
}

/** A PARTIAL write: only the columns the action changed. Never toListItem. */
export function lifecycleItem(fields) {
  const body = {};
  for (const [key, column] of Object.entries(TEXT_COLUMN)) {
    if (key in fields) body[column] = fields[key] === null || fields[key] === undefined ? '' : String(fields[key]);
  }
  if ('statusChangedOn' in fields) body.StatusChangedOn = new Date(fields.statusChangedOn).toISOString();
  if ('manualFields' in fields) body.ManualFields = fields.manualFields.join('\n');
  return body;
}

/** The owner-history writes a plan implies, in order: end the old, start the new. */
export function stintWrites({ close, open }) {
  const writes = [];
  if (close) {
    writes.push(close.stint.id !== null && close.stint.id !== undefined
      ? { id: close.stint.id, body: closeBody(close) }
      : {
        id: null,
        body: toAssignmentItem({
          ...close.stint, endedOn: close.endedOn, endReason: close.endReason, note: close.note ?? close.stint.note,
        }),
      });
  }
  if (open) writes.push({ id: null, body: toAssignmentItem(open) });
  return writes;
}

export async function writeStints({ siteUrl, token, digest, writes }) {
  for (let index = 0; index < writes.length; index += 1) {
    const write = writes[index];
    const path = write.id === null
      ? `${listPath(ASSIGNMENT_LIST_NAME)}/items`
      : `${listPath(ASSIGNMENT_LIST_NAME)}/items(${write.id})`;
    let response;
    try {
      response = await withRetry(() => spFetch(siteUrl, path, {
        token, digest, method: 'POST', accept: ITEM_ACCEPT, body: write.body,
        ...(write.id === null ? {} : { headers: MERGE }),
      }));
    } catch (thrown) {
      // Offline: spFetch throws instead of answering. Same rule as below.
      thrown.remaining = writes.slice(index);
      throw thrown;
    }
    if (!response.ok) {
      const failure = new Error(`Could not record the owner history (${response.status})`);
      // Only what has not landed is retried: re-posting an ended legacy stint
      // that already went in would put the same person in the history twice.
      failure.remaining = writes.slice(index);
      throw failure;
    }
  }
}

/**
 * Re-read, plan, write. The device row goes first and is the only write that
 * fails the action: a machine whose status moved but whose history did not is
 * recoverable from `pendingStints`; the reverse would be history describing a
 * move that never happened.
 */
export async function performLifecycle({
  siteUrl, token, deviceId, action, input, expectedStatus, recordedBy = '', now = Date.now(),
}) {
  const digest = (await ensureProvisioned(siteUrl, token)) ?? await getFormDigest(siteUrl, token);
  const device = await readDevice(siteUrl, token, deviceId);
  if (!device) throw new Error('That machine is no longer in the register.');
  const stints = await readAssignments(siteUrl, token, { deviceId });

  const plan = planLifecycle({ device, stints, action, input, expectedStatus, now, recordedBy });
  if (plan.refusal) throw new Error(plan.refusal);

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${device.id})`, {
    token, digest, method: 'POST', accept: ITEM_ACCEPT, body: lifecycleItem(plan.fields), headers: MERGE,
  }));
  if (!response.ok) throw new Error(`Could not save the change (${response.status}): ${await response.text()}`);

  let pendingStints = [];
  try {
    await writeStints({ siteUrl, token, digest, writes: stintWrites(plan) });
  } catch (failure) {
    pendingStints = failure.remaining ?? stintWrites(plan);
  }

  // The move happened. A history line that would not write must not turn it
  // into an error that invites a second press.
  try {
    await logChanges(siteUrl, token, digest, device, plan.changes, recordedBy);
  } catch {
    return { plan, pendingStints, logFailed: true };
  }
  return { plan, pendingStints };
}

/** Retry only the history writes an action left behind. */
export async function retryStints({ siteUrl, token, writes }) {
  const digest = await getFormDigest(siteUrl, token);
  await writeStints({ siteUrl, token, digest, writes });
}
