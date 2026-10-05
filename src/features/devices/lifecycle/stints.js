import { inFleet, statusOf, SPARE, RETIRED } from './status.js';

const newestFirst = (a, b) => (b.assignedOn ?? 0) - (a.assignedOn ?? 0);

export function stintsOf(deviceId, stints = []) {
  return stints.filter((stint) => stint.deviceId === deviceId).sort(newestFirst);
}

export function openStintOf(deviceId, stints = []) {
  return stintsOf(deviceId, stints).find((stint) => stint.endedOn === null || stint.endedOn === undefined) ?? null;
}

/**
 * What the register already says about a machine nobody has written history
 * for: its current owner, since at least the day the row appeared. Never
 * written until something ends it -- nothing is written for a machine nobody
 * touches.
 */
export function legacyStint(device) {
  if (!device?.owner) return null;
  return {
    id: null,
    deviceId: device.id,
    computerName: device.computerName ?? null,
    owner: device.owner,
    location: device.location ?? null,
    department: device.department ?? null,
    assignedOn: device.createdOn ?? device.importedOn ?? device.scannedOn ?? null,
    assignedOnApprox: true,
    endedOn: null,
    endReason: null,
    note: null,
    recordedBy: null,
  };
}

/** The stint a lifecycle action would end, or null when nobody has the machine. */
export function currentStint(device, stints = []) {
  const open = openStintOf(device.id, stints);
  if (open) return open;
  if (!inFleet(device)) return null;
  // History exists and none of it is open: the register's owner is not backed
  // by a stint, and inventing one now would put a guess after real records.
  if (stintsOf(device.id, stints).length) return null;
  return legacyStint(device);
}

const gapLabel = (reason) => {
  if (reason === 'To stash') return 'In IT Stash';
  if (reason === 'Retired') return 'In the Graveyard';
  return 'Nobody recorded';
};

const nowLabel = (device) => {
  if (statusOf(device) === SPARE) return 'In IT Stash';
  if (statusOf(device) === RETIRED) return 'In the Graveyard';
  return 'Nobody recorded';
};

/** Owner history newest first, with the spells nobody had it shown as gaps. */
export function timelineOf(device, stints = []) {
  const own = stintsOf(device.id, stints);
  const list = own.length ? own : [legacyStint(device)].filter(Boolean);
  const entries = [];

  const latest = list[0];
  if (latest && latest.endedOn !== null && latest.endedOn !== undefined) {
    entries.push({ kind: 'gap', from: latest.endedOn, to: null, label: nowLabel(device) });
  }

  list.forEach((stint, index) => {
    entries.push({ kind: 'stint', ...stint, current: stint.endedOn === null || stint.endedOn === undefined });
    const older = list[index + 1];
    if (older?.endedOn != null && stint.assignedOn != null && stint.assignedOn > older.endedOn) {
      entries.push({ kind: 'gap', from: older.endedOn, to: stint.assignedOn, label: gapLabel(older.endReason) });
    }
  });

  return entries;
}
