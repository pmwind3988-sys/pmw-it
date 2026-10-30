import { cleanSerial } from '../devices/parse/placeholders.js';
import { statusOf, inFleet, SPARE, RETIRED } from '../devices/lifecycle/status.js';
import { ACTIONS } from '../devices/lifecycle/planLifecycle.js';

/**
 * Pure: when an asset-register row and a device-list machine are the SAME
 * physical computer, and what the device list has to do to agree with the
 * register.
 *
 * The two lists were built separately -- the register from boxes and labels,
 * the device list from scan reports -- so a laptop handed out at the till used
 * to stay "in the IT Stash" on the device list. The register is where a
 * handover is recorded, so the register leads and the device list follows.
 *
 * Same machine means, in order:
 *  1. the same serial (a Dell service tag is both the box's barcode and the
 *     BIOS serial the scan reads) -- any category;
 *  2. the row's label is the machine's computer name, now or before a rename
 *     (IT labels a laptop PMWL010 and names it PMWL010) -- computers only, so
 *     a monitor labelled like a PC is never taken for one.
 * A serial or a name that points at more than one machine links nothing:
 * guessing which of two would move the wrong one.
 */

const COMPUTERS = new Set(['Laptop', 'Desktop']);
const OUT = new Set(['Assigned', 'Borrowed']);
const text = (value) => String(value ?? '').trim();
const nameKey = (value) => text(value).toLowerCase();

const namesOf = (device) => [device?.computerName, ...(device?.previousNames ?? [])]
  .map(nameKey).filter(Boolean);

function only(list) {
  return list.length === 1 ? list[0] : null;
}

export function machineFor(asset, devices = []) {
  if (!asset) return null;
  const serial = cleanSerial(asset.serialNumber);
  if (serial) {
    const bySerial = devices.filter((device) => device?.id != null && cleanSerial(device.serialNumber) === serial);
    if (bySerial.length) {
      const device = only(bySerial);
      return device ? { device, by: 'serial' } : null;
    }
  }
  const label = nameKey(asset.assetTag);
  if (!label || !COMPUTERS.has(asset.category)) return null;
  const device = only(devices.filter((entry) => entry?.id != null && namesOf(entry).includes(label)));
  return device ? { device, by: 'name' } : null;
}

/** The register row that is this machine, if exactly one is. */
export function assetFor(device, assets = []) {
  if (!device) return null;
  const linked = assets.filter((asset) => machineFor(asset, [device])?.device === device);
  // Two rows both claiming one machine: the same rule as above, link neither.
  return only(linked);
}

/** Who the register says has it, or null when it is not out. */
export function holderOf(asset) {
  if (!asset || !OUT.has(asset.status)) return null;
  const name = text(asset.assignedTo);
  return name ? { name, email: text(asset.assignedToEmail) } : null;
}

/**
 * The device-list move that makes the machine agree with the register, or
 * null when they already agree or the register says nothing either way.
 *
 * - Out with somebody in the register: the machine is theirs (Assign from the
 *   stash or the Graveyard, Change owner if somebody else has it).
 * - In stock in the register, and only straight after a RETURN
 *   (`justReturned`): a machine somebody has goes to the IT Stash. Otherwise
 *   "In stock" proves nothing -- a laptop counted into the register as
 *   in stock may well be on a desk, and the device list knows better.
 * Repair, retirement and disposal in the register move nothing: those are
 * decisions about a machine, taken on the machine's own page.
 */
export function machineMove(asset, device, { justReturned = false } = {}) {
  if (!asset || !device) return null;
  const status = statusOf(device);
  const holder = holderOf(asset);
  if (holder) {
    const keep = { location: device.location ?? null, department: device.department ?? null };
    if (status === SPARE || status === RETIRED) {
      return { action: ACTIONS.ASSIGN, expectedStatus: status, input: { owner: holder.name, ...keep }, holder };
    }
    if (nameKey(device.owner) === nameKey(holder.name)) return null;
    return { action: ACTIONS.CHANGE_OWNER, expectedStatus: status, input: { owner: holder.name, ...keep }, holder };
  }
  if (justReturned && asset.status === 'In stock' && inFleet(device)) {
    return { action: ACTIONS.TO_STASH, expectedStatus: status, input: { endReason: 'Returned' }, holder: null };
  }
  return null;
}

/** In words, for a banner: what the register says and what pressing does. */
export function describeMove(move, device) {
  if (!move) return '';
  if (move.action === ACTIONS.TO_STASH) {
    return `The asset register has it back in stock, but here it is still with ${text(device?.owner) || 'somebody'}.`;
  }
  const here = statusOf(device) === SPARE ? 'here it is in the IT Stash'
    : statusOf(device) === RETIRED ? 'here it is retired'
      : `here it is with ${text(device?.owner) || 'nobody named'}`;
  return `The asset register has it with ${move.holder.name}, but ${here}.`;
}

/** Every machine the register disagrees with, for the To do view. */
export function disagreements(assets = [], devices = []) {
  const out = [];
  for (const asset of assets) {
    const link = machineFor(asset, devices);
    if (!link || assetFor(link.device, assets) !== asset) continue;
    const move = machineMove(asset, link.device);
    if (move) out.push({ asset, device: link.device, move });
  }
  return out;
}
