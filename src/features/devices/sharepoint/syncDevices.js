import { spFetch, listPath, ITEM_ACCEPT } from './spClient.js';
import { provisionLists } from './provisionLists.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './deviceSchema.js';
import { readAllDevices } from './readDevices.js';
import { diffDevice, indexByName } from './diffDevice.js';
import { runPool, withRetry } from './writePool.js';
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';

/**
 * Pure: decides what to write. Kept separate from syncDevices so that every
 * insert/update/skip decision is testable without a token or a network.
 */
export function planSync(incoming, existingIndex) {
  const inserts = [];
  const updates = [];
  const changeRows = [];

  for (const device of incoming) {
    const key = String(device.computerName ?? '').toLowerCase();
    const existing = existingIndex.get(key);
    const body = toListItem(device);

    if (!existing) {
      inserts.push({ computerName: device.computerName, body });
      continue;
    }

    const changes = diffDevice(existing, device);
    if (!changes.length) continue;

    updates.push({ computerName: device.computerName, id: existing.id, body });
    for (const change of changes) {
      changeRows.push({ computerName: device.computerName, ...change });
    }
  }

  return { inserts, updates, changeRows };
}

const itemPath = (listName) => `${listPath(listName)}/items`;

/**
 * Progress is reported as `{ phase, done, total }` rather than a bare pair,
 * because the row writes are the short part. A first run spends over a minute
 * provisioning ~70 columns before the first row moves, and a bar that sits at
 * "0 of 3" throughout is indistinguishable from a hang.
 */
export async function syncDevices({ siteUrl, token, devices, changedBy, onProgress }) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  // Provisioning runs first and throws on failure: a half-created list would
  // fail every row with the same unhelpful message.
  report('provisioning');
  const digest = await provisionLists(siteUrl, token, {
    onProgress: (done, total) => report('provisioning', done, total),
  });

  report('reading');
  const existing = await readAllDevices(siteUrl, token);
  const plan = planSync(devices, indexByName(existing));

  const post = (path, body) =>
    withRetry(() =>
      spFetch(siteUrl, path, {
        token, digest, method: 'POST', body, accept: ITEM_ACCEPT,
      }));

  const work = [
    ...plan.inserts.map((entry) => ({ ...entry, action: 'insert' })),
    ...plan.updates.map((entry) => ({ ...entry, action: 'update' })),
  ];

  report('writing', 0, work.length);
  const results = await runPool(
    work,
    async (entry) => {
      const response = entry.action === 'insert'
        ? await post(itemPath(DEVICE_LIST_NAME), entry.body)
        : await withRetry(() =>
          spFetch(siteUrl, `${itemPath(DEVICE_LIST_NAME)}(${entry.id})`, {
            token,
            digest,
            method: 'POST',
            body: entry.body,
            accept: ITEM_ACCEPT,
            // A SharePoint update is a POST wearing these two headers.
            headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
          }));

      if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
      return entry.action;
    },
    { concurrency: 4, onProgress: (done, total) => report('writing', done, total) },
  );

  const changedOn = Date.now();
  if (plan.changeRows.length) report('logging', 0, plan.changeRows.length);
  const changeResults = await runPool(plan.changeRows, async (row) => {
    const response = await post(itemPath(CHANGE_LIST_NAME), {
      Title: row.computerName,
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
    changeCount: plan.changeRows.length,
    changeFailures: changeResults.filter((result) => result.error).length,
    unchanged: devices.length - work.length,
  };
}
