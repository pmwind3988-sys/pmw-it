import {
  statusOf, IN_USE, IN_REPAIR, SPARE, RETIRED,
} from './status.js';
import { currentStint } from './stints.js';
import { cleanLocation } from '../map/locations.js';

/**
 * What one action on one machine writes. Pure: the caller re-reads the
 * machine and its stints immediately before calling, and performs the result
 * in order -- device row, then stints, then change log.
 */
export const ACTIONS = {
  CHANGE_OWNER: 'changeOwner',
  TO_REPAIR: 'toRepair',
  BACK_IN_USE: 'backInUse',
  TO_STASH: 'toStash',
  RETIRE: 'retire',
  ASSIGN: 'assign',
};

const ALLOWED = {
  [IN_USE]: [ACTIONS.CHANGE_OWNER, ACTIONS.TO_REPAIR, ACTIONS.TO_STASH, ACTIONS.RETIRE],
  [IN_REPAIR]: [ACTIONS.CHANGE_OWNER, ACTIONS.BACK_IN_USE, ACTIONS.TO_STASH, ACTIONS.RETIRE],
  [SPARE]: [ACTIONS.ASSIGN, ACTIONS.RETIRE],
  [RETIRED]: [ACTIONS.ASSIGN],
};

export function actionsFor(device) {
  return ALLOWED[statusOf(device)];
}

const WHERE = { [SPARE]: 'IT Stash', [RETIRED]: 'the Graveyard', [IN_REPAIR]: 'repair', [IN_USE]: 'use' };
const text = (value) => (value === null || value === undefined ? '' : String(value).trim());
const same = (a, b) => text(a).toUpperCase() === text(b).toUpperCase();

export function planLifecycle({
  device, stints = [], action, input = {}, expectedStatus, now = Date.now(), recordedBy = '',
}) {
  const status = statusOf(device);

  if (expectedStatus && expectedStatus !== status) {
    return { refusal: `This machine was just moved to ${WHERE[status]} by someone else. Refresh to see it as it is now.` };
  }
  if (!ALLOWED[status].includes(action)) {
    return { refusal: `A machine that is ${status.toLowerCase()} cannot do that.` };
  }

  const on = typeof input.on === 'number' ? input.on : now;
  const note = text(input.note) || null;
  const stint = currentStint(device, stints);
  const manual = new Set(device.manualFields ?? []);
  const fields = {};
  const changes = [];
  let close = null;
  let open = null;

  const set = (key, value) => {
    const before = text(device[key]);
    const after = text(value);
    if (before === after) return;
    fields[key] = value;
    let changeType = 'Updated';
    if (!before) changeType = 'Added';
    else if (!after) changeType = 'Removed';
    changes.push({ fieldName: key, oldValue: before, newValue: after, changeType });
  };
  const setStatus = (next) => {
    set('status', next);
    fields.statusChangedOn = on;
  };
  const holder = () => {
    const owner = text(input.owner);
    if (!owner) return null;
    return { owner, location: cleanLocation(input.location), department: text(input.department) || null };
  };
  const take = (who) => {
    set('owner', who.owner);
    set('location', who.location);
    set('department', who.department);
    fields.ownerSource = 'Manual';
    ['owner', 'location', 'department'].forEach((key) => manual.add(key));
  };
  const release = () => {
    set('owner', null);
    set('department', null);
    // Blank and handed back to the scan: a later scan naming somebody is how a
    // spare machine comes back into use without anyone pressing anything.
    manual.delete('owner');
    manual.delete('department');
  };
  const end = (reason) => {
    if (stint) close = { stint, endedOn: on, endReason: reason, note };
  };
  const start = (who) => {
    open = {
      deviceId: device.id,
      computerName: device.computerName ?? null,
      ...who,
      assignedOn: on,
      assignedOnApprox: false,
      endedOn: null,
      endReason: null,
      // A note explains why a stint ENDED when one ends; otherwise it belongs
      // to the one starting.
      note: close ? null : note,
      recordedBy,
    };
  };

  switch (action) {
    case ACTIONS.CHANGE_OWNER: {
      const who = holder();
      if (!who) return { refusal: 'Type who the machine is going to.' };
      if (same(who.owner, device.owner) && same(who.location, device.location) && same(who.department, device.department)) {
        return { refusal: `${who.owner} already has this machine.` };
      }
      end('Reassigned');
      take(who);
      start(who);
      break;
    }
    case ACTIONS.TO_REPAIR:
      setStatus(IN_REPAIR);
      break;
    case ACTIONS.BACK_IN_USE:
      setStatus(IN_USE);
      break;
    case ACTIONS.TO_STASH:
      end(input.endReason ?? 'To stash');
      release();
      setStatus(SPARE);
      break;
    case ACTIONS.RETIRE:
      end(input.endReason ?? 'Retired');
      release();
      setStatus(RETIRED);
      break;
    case ACTIONS.ASSIGN: {
      const who = holder();
      if (!who) return { refusal: 'Type who the machine is going to.' };
      take(who);
      setStatus(IN_USE);
      start(who);
      break;
    }
    default:
      return { refusal: 'That is not something a machine can do.' };
  }

  fields.manualFields = [...manual];
  return { fields, changes, close, open };
}
