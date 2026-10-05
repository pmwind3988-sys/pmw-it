import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionLists } from './provisionLists.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './deviceSchema.js';
import { readAllDevices } from './readDevices.js';
import { readAssignments } from './readAssignments.js';
import { diffDevice } from './diffDevice.js';
import { stintWrites, writeStints, lifecycleItem } from './writeLifecycle.js';
import { runPool, withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';
import { matchIncoming, MATCH } from '../lifecycle/matchIncoming.js';
import { replacementsFor, ANSWERS } from '../lifecycle/replacements.js';
import { planLifecycle, ACTIONS } from '../lifecycle/planLifecycle.js';
import { currentStint } from '../lifecycle/stints.js';
import { inFleet, IN_USE } from '../lifecycle/status.js';

/**
 * A field somebody corrected by hand outranks what the scan file says about
 * it. Without this the Edit button in the register would be a trap: the next
 * import of the same unchanged file would silently undo the correction.
 *
 * Only the named fields are held back, and the list itself is carried forward
 * so that updating anything else does not wipe it.
 */
export function applyManualOverrides(incoming, existing) {
  const manual = existing?.manualFields;
  if (!Array.isArray(manual) || !manual.length) return incoming;

  const merged = { ...incoming, manualFields: manual };
  for (const field of manual) {
    // A name left over from a field that no longer exists is skipped rather
    // than writing `undefined` over a real value.
    if (field in existing) merged[field] = existing[field];
  }
  return merged;
}

const ownerKey = (owner) => String(owner ?? '').trim().replace(/\s+/g, ' ').toLowerCase();

/**
 * Pure: every decision an import makes, in one place, so all of it is
 * testable without a token. Replaces planSync, which matched by name only.
 */
export function planImport(incoming, existing, {
  stints = [], answers = {}, now = Date.now(), recordedBy = '',
} = {}) {
  const matches = matchIncoming(incoming, existing);
  const prompts = replacementsFor(matches, existing);
  const inserts = [];
  const updates = [];
  const changeRows = [];
  const stintOps = [];
  const retirements = [];
  const skipped = [];

  const replacing = new Set(prompts
    .filter((prompt) => answers[prompt.key] && answers[prompt.key] !== ANSWERS.KEEP)
    .map((prompt) => prompt.sourceFileName));

  for (const match of matches) {
    const { device, existing: row, kind } = match;
    const on = device.scannedOn ?? now;
    const holder = { owner: device.owner, location: device.location ?? null, department: device.department ?? null };

    if (kind === MATCH.DUPLICATE_SERIAL) {
      skipped.push({ computerName: device.computerName, reason: match.note });
      continue;
    }

    if (!row) {
      inserts.push({
        computerName: device.computerName,
        body: toListItem({ ...device, status: IN_USE, statusChangedOn: on }),
      });
      if (device.owner) {
        stintOps.push({
          computerName: device.computerName,
          deviceId: null,
          close: null,
          open: {
            deviceId: null, computerName: device.computerName, ...holder,
            assignedOn: on, assignedOnApprox: !replacing.has(device.sourceFileName),
            endedOn: null, endReason: null, note: null, recordedBy,
          },
        });
      }
      continue;
    }

    // An older report carries no location or serial; it must not erase them.
    const carried = {
      ...device,
      location: device.location ?? row.location ?? null,
      serialNumber: device.serialNumber ?? row.serialNumber ?? null,
    };
    const resolved = applyManualOverrides(carried, row);
    const wasActive = inFleet(row);
    resolved.status = wasActive ? (row.status ?? null) : IN_USE;
    if (!wasActive) resolved.statusChangedOn = on;

    const changes = diffDevice(row, resolved);
    if (changes.length) {
      updates.push({ computerName: device.computerName, id: row.id, body: toListItem(resolved) });
      for (const change of changes) changeRows.push({ computerName: device.computerName, deviceId: row.id, ...change });
    }

    const who = { owner: resolved.owner, location: resolved.location ?? null, department: resolved.department ?? null };
    const startFor = () => ({
      deviceId: row.id, computerName: resolved.computerName, ...who,
      assignedOn: on, assignedOnApprox: false, endedOn: null, endReason: null, note: null, recordedBy,
    });

    if (wasActive && resolved.owner && ownerKey(row.owner) !== ownerKey(resolved.owner)) {
      const stint = currentStint(row, stints);
      stintOps.push({
        computerName: resolved.computerName,
        deviceId: row.id,
        close: stint ? { stint, endedOn: on, endReason: 'Reassigned', note: null } : null,
        open: startFor(),
      });
    } else if (!wasActive && resolved.owner) {
      stintOps.push({ computerName: resolved.computerName, deviceId: row.id, close: null, open: startFor() });
    }
  }

  const handled = new Set();
  for (const prompt of prompts) {
    const answer = answers[prompt.key];
    if (!answer || answer === ANSWERS.KEEP || handled.has(prompt.old.id)) continue;
    handled.add(prompt.old.id);

    const old = existing.find((row) => row.id === prompt.old.id);
    const incomingScan = incoming.find((device) => device.sourceFileName === prompt.sourceFileName);
    const plan = planLifecycle({
      device: old,
      stints,
      action: answer === ANSWERS.RETIRE ? ACTIONS.RETIRE : ACTIONS.TO_STASH,
      input: { endReason: 'Replaced', on: incomingScan?.scannedOn ?? now },
      now,
      recordedBy,
    });
    if (plan.refusal) continue;

    retirements.push({ id: old.id, computerName: old.computerName, body: lifecycleItem(plan.fields), changes: plan.changes });
    stintOps.push({ computerName: old.computerName, deviceId: old.id, close: plan.close, open: null });
  }

  return { matches, prompts, inserts, updates, changeRows, stintOps, retirements, skipped };
}

const itemPath = (listName) => `${listPath(listName)}/items`;

/**
 * Progress is reported as `{ phase, done, total }` rather than a bare pair,
 * because the row writes are the short part. A first run spends over a minute
 * provisioning ~70 columns before the first row moves, and a bar that sits at
 * "0 of 3" throughout is indistinguishable from a hang.
 */
export async function syncDevices({
  siteUrl, token, devices, changedBy, onProgress, answers = {},
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  report('provisioning');
  const digest = await provisionLists(siteUrl, token, {
    onProgress: (done, total) => report('provisioning', done, total),
  });

  report('reading');
  const existing = await readAllDevices(siteUrl, token);
  const stints = await readAssignments(siteUrl, token);
  const plan = planImport(devices, existing, { stints, answers, recordedBy: changedBy ?? '' });

  const post = (path, body) =>
    withRetry(() => spFetch(siteUrl, path, { token, digest, method: 'POST', body, accept: ITEM_ACCEPT }));
  const merge = (id, body) =>
    withRetry(() => spFetch(siteUrl, `${itemPath(DEVICE_LIST_NAME)}(${id})`, {
      token, digest, method: 'POST', body, accept: ITEM_ACCEPT,
      // A SharePoint update is a POST wearing these two headers.
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    }));

  const work = [
    ...plan.inserts.map((entry) => ({ ...entry, action: 'insert' })),
    ...plan.updates.map((entry) => ({ ...entry, action: 'update' })),
    ...plan.retirements.map((entry) => ({ ...entry, action: 'retire' })),
  ];

  report('writing', 0, work.length);
  const results = await runPool(
    work,
    async (entry) => {
      const response = entry.action === 'insert'
        ? await post(itemPath(DEVICE_LIST_NAME), entry.body)
        : await merge(entry.id, entry.body);
      if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
      if (entry.action !== 'insert') return { id: entry.id };
      const created = await response.json().catch(() => ({}));
      return { id: created?.Id ?? created?.d?.Id ?? null };
    },
    { concurrency: 4, onProgress: (done, total) => report('writing', done, total) },
  );

  // A new machine's first stint can only point at it once SharePoint has
  // given the row an id; a machine whose row failed gets no history at all.
  const idOf = new Map();
  const landed = new Set();
  results.forEach((result, index) => {
    if (result.error) return;
    landed.add(work[index].computerName);
    if (work[index].action === 'insert') idOf.set(work[index].computerName, result.value?.id ?? null);
  });

  let stintsWritten = 0;
  let stintFailures = 0;
  for (const op of plan.stintOps) {
    if (!landed.has(op.computerName)) continue;
    const deviceId = op.deviceId ?? idOf.get(op.computerName);
    if (deviceId === null || deviceId === undefined) {
      stintFailures += 1;
      continue;
    }
    const writes = stintWrites({ close: op.close, open: op.open ? { ...op.open, deviceId } : null });
    try {
      await writeStints({ siteUrl, token, digest, writes });
      stintsWritten += writes.length;
    } catch {
      stintFailures += 1;
    }
  }

  const retirementRows = plan.retirements.flatMap((entry) => entry.changes.map((change) => ({
    computerName: entry.computerName, deviceId: entry.id, ...change,
  })));
  const changeRows = [...plan.changeRows, ...retirementRows];

  const changedOn = Date.now();
  if (changeRows.length) report('logging', 0, changeRows.length);
  const changeResults = await runPool(changeRows, async (row) => {
    const response = await post(itemPath(CHANGE_LIST_NAME), {
      Title: row.computerName,
      DeviceId: row.deviceId ?? idOf.get(row.computerName) ?? null,
      FieldName: row.fieldName,
      OldValue: row.oldValue,
      NewValue: row.newValue,
      ChangeType: row.changeType,
      ChangedOn: new Date(changedOn).toISOString(),
      ChangedOnMYT: formatMYT(changedOn, 'datetime12'),
      ChangedBy: changedBy ?? '',
    });
    if (!response.ok) throw new Error(String(response.status));
    return true;
  }, {
    concurrency: 4,
    onProgress: (done, total) => report('logging', done, total),
  });

  return {
    results: results.map((result, index) => ({
      computerName: work[index].computerName,
      action: work[index].action,
      error: result.error ? result.error.message : null,
    })),
    changeCount: changeRows.length,
    changeFailures: changeResults.filter((result) => result.error).length,
    unchanged: devices.length - plan.inserts.length - plan.updates.length - plan.skipped.length,
    stintsWritten,
    stintFailures,
    skipped: plan.skipped,
  };
}
