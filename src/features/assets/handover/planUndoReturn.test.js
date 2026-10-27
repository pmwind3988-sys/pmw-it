import { describe, it, expect } from 'vitest';
import { lastReturnOf, planUndoReturn } from './planUndoReturn.js';
import { planReturn, returnEverything } from './planReturn.js';
import { HANDOVER_STATUS, HANDOVER_KIND } from './availability.js';

const laptop = {
  id: 1, assetKey: 'serial:DELL|CN0ABC', trackingMode: 'Tracked', quantity: 1, quantityOut: 1,
  status: 'Assigned', assignedTo: 'Thamil Selvi', assignedToEmail: 'thamil@pmw.com',
};
const cables = { id: 2, assetKey: 'bulk:CABLE||HDMI', trackingMode: 'Bulk', quantity: 20, quantityOut: 5 };

const laptopOut = {
  id: 11, assetId: 1, assetKey: laptop.assetKey, personEmail: 'thamil@pmw.com',
  personName: 'Thamil Selvi', quantity: 1, returnedQuantity: 0,
  handoverStatus: HANDOVER_STATUS.OUT, kind: HANDOVER_KIND.ISSUED, issuedOn: 1000,
};
const cablesOut = {
  id: 12, assetId: 2, assetKey: cables.assetKey, personEmail: 'thamil@pmw.com',
  personName: 'Thamil Selvi', quantity: 3, returnedQuantity: 0,
  handoverStatus: HANDOVER_STATUS.OUT, kind: HANDOVER_KIND.ISSUED, issuedOn: 2000,
};

/** Apply a plan's bodies to the rows, the way SharePoint would. */
const apply = (rows, updates) => rows.map((row) => {
  const update = updates.find((entry) => entry.id === row.id);
  return update ? { ...row, ...update.body } : row;
});

/** Thamil's state right after "Return everything". */
function afterReturnEverything() {
  const handovers = [laptopOut, cablesOut];
  const register = [laptop, cables];
  const plan = planReturn(returnEverything(handovers, 'Good'), handovers, register, {
    returnedOn: 5000, returnedBy: 'it@pmw',
  });
  return {
    handovers: apply(handovers, plan.handoverUpdates),
    register: apply(register, plan.assetUpdates),
  };
}

describe('finding the last return', () => {
  it('takes every row stamped with the latest return time', () => {
    const { handovers } = afterReturnEverything();
    const older = { ...laptopOut, id: 9, returnedQuantity: 1, returnedOn: 10 };
    const found = lastReturnOf([...handovers, older], 'THAMIL@pmw.com');
    expect(found.returnedOn).toBe(5000);
    expect(found.rows.map((row) => row.id)).toEqual([11, 12]);
  });

  it('finds nothing for somebody who never returned anything', () => {
    expect(lastReturnOf([laptopOut], 'thamil@pmw.com').rows).toEqual([]);
  });
});

describe('undoing "Return everything"', () => {
  it('puts the handover rows and the register back as they were', () => {
    const { handovers, register } = afterReturnEverything();
    const plan = planUndoReturn(lastReturnOf(handovers, 'thamil@pmw.com').rows, register);

    expect(plan.blocked).toEqual([]);
    for (const update of plan.handoverUpdates) {
      expect(update.body).toEqual({
        returnedQuantity: 0,
        handoverStatus: HANDOVER_STATUS.OUT,
        returnedOn: null,
        returnedBy: '',
        returnCondition: null,
        returnSignature: '',
      });
    }

    const restored = apply(register, plan.assetUpdates);
    expect(restored[0]).toMatchObject({
      quantityOut: 1, status: 'Assigned', handoverKind: 'Issued',
      assignedTo: 'Thamil Selvi', assignedToEmail: 'thamil@pmw.com', assignedOn: 1000, dueOn: null,
    });
    expect(restored[1].quantityOut).toBe(5);
  });

  it('refuses a laptop somebody else has been given since', () => {
    const { handovers, register } = afterReturnEverything();
    const reissued = register.map((asset) => (asset.id === 1
      ? { ...asset, quantityOut: 1, assignedTo: 'Amir', assignedToEmail: 'amir@pmw.com' }
      : asset));
    const plan = planUndoReturn(lastReturnOf(handovers, 'thamil@pmw.com').rows, reissued);

    expect(plan.blocked).toHaveLength(1);
    expect(plan.blocked[0].reason).toMatch(/Amir/);
    expect(plan.handoverUpdates.map((update) => update.id)).toEqual([12]);
  });

  it('refuses cables that are no longer on the shelf', () => {
    const { handovers, register } = afterReturnEverything();
    const empty = register.map((asset) => (asset.id === 2 ? { ...asset, quantityOut: 19 } : asset));
    const plan = planUndoReturn(lastReturnOf(handovers, 'thamil@pmw.com').rows, empty);
    expect(plan.blocked.map((entry) => entry.row.id)).toEqual([12]);
  });

  it('adds two lines of one box together against one register row', () => {
    const rows = [
      { ...cablesOut, id: 21, returnedQuantity: 2, returnedOn: 1 },
      { ...cablesOut, id: 22, returnedQuantity: 1, returnedOn: 1 },
    ];
    const plan = planUndoReturn(rows, [{ ...cables, quantityOut: 0 }]);
    expect(plan.assetUpdates).toHaveLength(1);
    expect(plan.assetUpdates[0].body.quantityOut).toBe(3);
  });

  it('puts a borrowed item back with its due date', () => {
    const row = {
      ...laptopOut, kind: HANDOVER_KIND.BORROWED, dueOn: 7000, returnedQuantity: 1, returnedOn: 5,
    };
    const shelf = { ...laptop, quantityOut: 0, assignedTo: '', assignedToEmail: '', status: 'In stock' };
    const plan = planUndoReturn([row], [shelf]);
    expect(plan.assetUpdates[0].body).toMatchObject({ status: 'Borrowed', dueOn: 7000 });
  });
});
