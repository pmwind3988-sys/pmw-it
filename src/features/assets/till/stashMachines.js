import { statusOf, SPARE } from '../../devices/lifecycle/status.js';
import { normaliseCode } from '../identity.js';

/**
 * Machines waiting in the IT Stash (`Spare` in the device list), handed out
 * from the till like anything else.
 *
 * A machine is not an asset-register row: it lives in the device list, with
 * its own owner history. So it rides on the receipt beside the basket, and
 * checkout gives it to the person in the device list (the same "assign" the
 * device page does) and puts it on the checklist they sign — ticked as a
 * Laptop or Desktop, its serial in the serial box.
 */

const LIMIT = 6;
const text = (value) => String(value ?? '').trim();

export function inStash(devices = []) {
  return devices
    .filter((device) => device?.id != null && statusOf(device) === SPARE)
    .sort((a, b) => (Number(b.deviceType === 'Laptop') - Number(a.deviceType === 'Laptop'))
      || text(a.computerName).localeCompare(text(b.computerName)));
}

/** What the machine is called on a receipt: its model, then its name. */
export function machineTitle(device) {
  const model = text(device?.computerModel);
  const name = text(device?.computerName);
  if (model && name) return `${model} (${name})`;
  return model || name || 'Machine';
}

export function machineSub(device) {
  return [
    text(device?.serialNumber) && `S/N ${text(device.serialNumber)}`,
    'IT Stash',
  ].filter(Boolean).join(' · ');
}

/** A scanned code that is a stash machine's serial or computer name. */
export function findInStash(devices, raw, picked = []) {
  const code = normaliseCode(raw);
  if (!code) return null;
  const taken = new Set(picked.map((device) => device.id));
  return inStash(devices).find((device) => !taken.has(device.id)
    && (normaliseCode(device.serialNumber) === code || normaliseCode(device.computerName) === code)) ?? null;
}

/** Type-to-find over the stash, shaped like the till's other results. */
export function searchStash(devices, query, picked = []) {
  const needle = text(query).toLowerCase();
  if (needle.length < 2) return [];
  const taken = new Set(picked.map((device) => device.id));
  return inStash(devices)
    .filter((device) => !taken.has(device.id))
    .filter((device) => [device.computerName, device.computerModel, device.serialNumber, device.cpuModel]
      .some((value) => String(value ?? '').toLowerCase().includes(needle)))
    .slice(0, LIMIT)
    .map((device) => ({
      id: `machine:${device.id}`,
      kind: 'machine',
      name: machineTitle(device),
      sub: machineSub(device),
      device,
    }));
}

/** A machine as a `checklistFromTill` row. */
export function machineRow(device) {
  const category = device?.deviceType === 'Laptop' || device?.deviceType === 'Desktop'
    ? device.deviceType
    : 'Computer';
  return { category, name: machineTitle(device), serial: text(device?.serialNumber), qty: 1 };
}
