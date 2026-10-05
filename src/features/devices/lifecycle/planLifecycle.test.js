import { describe, it, expect } from 'vitest';
import { ACTIONS, actionsFor, planLifecycle } from './planLifecycle.js';

const NOW = Date.UTC(2026, 9, 5, 4);
const machine = (over) => ({
  id: 4, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  status: null, manualFields: [], createdOn: Date.UTC(2026, 7, 21), ...over,
});
const open = { id: 30, deviceId: 4, owner: 'Amir', assignedOn: Date.UTC(2026, 1, 12), endedOn: null };
const plan = (over) => planLifecycle({ device: machine(), stints: [open], now: NOW, recordedBy: 'it', ...over });

describe('actionsFor', () => {
  it('offers what each status allows', () => {
    expect(actionsFor(machine())).toEqual(['changeOwner', 'toRepair', 'toStash', 'retire']);
    expect(actionsFor(machine({ status: 'In repair' }))).toEqual(['changeOwner', 'backInUse', 'toStash', 'retire']);
    expect(actionsFor(machine({ status: 'Spare' }))).toEqual(['assign', 'retire']);
    expect(actionsFor(machine({ status: 'Retired' }))).toEqual(['assign']);
  });
});

describe('planLifecycle — change owner', () => {
  const result = plan({
    action: ACTIONS.CHANGE_OWNER,
    input: { owner: 'Aisyah', location: 'f1', department: 'ENGINEERING', note: 'Amir got a workstation' },
  });

  it('ends the current stint as Reassigned, with the note', () => {
    expect(result.close).toEqual({ stint: open, endedOn: NOW, endReason: 'Reassigned', note: 'Amir got a workstation' });
  });

  it('opens one for the new holder', () => {
    expect(result.open).toMatchObject({
      deviceId: 4, computerName: 'AMIR-HP', owner: 'Aisyah', location: 'F1', department: 'ENGINEERING',
      assignedOn: NOW, assignedOnApprox: false, endedOn: null, note: null, recordedBy: 'it',
    });
  });

  it('marks the holder as set by hand so an import leaves it alone', () => {
    expect(result.fields).toMatchObject({ owner: 'Aisyah', ownerSource: 'Manual' });
    expect(result.fields.manualFields.sort()).toEqual(['department', 'location', 'owner']);
    expect(result.changes).toEqual([
      { fieldName: 'owner', oldValue: 'Amir', newValue: 'Aisyah', changeType: 'Updated' },
    ]);
  });

  it('refuses a change to the same holder', () => {
    expect(plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: 'amir', location: 'F1', department: 'engineering' } }).refusal)
      .toMatch(/already has/);
  });

  it('refuses without a name', () => {
    expect(plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: '  ' } }).refusal).toMatch(/Type who/);
  });

  it('treats a move to another location as a new stint', () => {
    const moved = plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: 'Amir', location: 'F3', department: 'ENGINEERING' } });
    expect(moved.refusal).toBeUndefined();
    expect(moved.changes).toEqual([{ fieldName: 'location', oldValue: 'F1', newValue: 'F3', changeType: 'Updated' }]);
  });
});

describe('planLifecycle — repair', () => {
  it('keeps the stint open: it is still their laptop', () => {
    const result = plan({ action: ACTIONS.TO_REPAIR });
    expect(result.close).toBeNull();
    expect(result.open).toBeNull();
    expect(result.fields).toMatchObject({ status: 'In repair', statusChangedOn: NOW });
  });

  it('comes back into use', () => {
    expect(plan({ device: machine({ status: 'In repair' }), action: ACTIONS.BACK_IN_USE }).fields.status).toBe('In use');
  });
});

describe('planLifecycle — stash and retire', () => {
  it('ends the stint and clears the holder going to the stash', () => {
    const result = plan({ device: machine({ manualFields: ['owner', 'deviceType'] }), action: ACTIONS.TO_STASH });
    expect(result.close.endReason).toBe('To stash');
    expect(result.fields).toMatchObject({ status: 'Spare', owner: null, department: null });
    expect(result.fields.manualFields).toEqual(['deviceType']);
  });

  it('records a replacement as Replaced, not To stash', () => {
    expect(plan({ action: ACTIONS.TO_STASH, input: { endReason: 'Replaced' } }).close.endReason).toBe('Replaced');
  });

  it('retires, ending the legacy stint of a machine with no history', () => {
    const result = plan({ stints: [], action: ACTIONS.RETIRE });
    expect(result.close.stint).toMatchObject({ id: null, owner: 'Amir', assignedOnApprox: true });
    expect(result.close.endReason).toBe('Retired');
    expect(result.fields.status).toBe('Retired');
  });

  it('brings a retired machine back for somebody', () => {
    const result = plan({
      device: machine({ status: 'Retired', owner: null, department: null }),
      stints: [],
      action: ACTIONS.ASSIGN,
      input: { owner: 'Hafizah', location: 'PML', department: 'SALES', note: 'loan' },
    });
    expect(result.close).toBeNull();
    expect(result.open).toMatchObject({ owner: 'Hafizah', location: 'PML', note: 'loan' });
    expect(result.fields.status).toBe('In use');
  });
});

describe('planLifecycle — refusals', () => {
  it('refuses when somebody else moved the machine first', () => {
    expect(plan({ device: machine({ status: 'Spare' }), action: ACTIONS.ASSIGN, input: { owner: 'X' }, expectedStatus: 'In use' }).refusal)
      .toMatch(/just moved to IT Stash by someone else/);
  });

  it('refuses an action the status does not allow', () => {
    expect(plan({ action: ACTIONS.ASSIGN, input: { owner: 'X' } }).refusal).toMatch(/cannot/);
  });

  it('uses the effective date given', () => {
    const on = Date.UTC(2026, 8, 30, 4);
    expect(plan({ action: ACTIONS.TO_STASH, input: { on } }).close.endedOn).toBe(on);
  });
});
