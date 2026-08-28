import { describe, it, expect } from 'vitest';
import { planReturn, returnEverything } from './planReturn.js';
import { HANDOVER_STATUS, HANDOVER_KIND } from './availability.js';

const laptop = { id: 1, assetKey: 'serial:DELL|CN0ABC', trackingMode: 'Tracked', quantity: 1, quantityOut: 1 };
const cables = { id: 2, assetKey: 'bulk:CABLE||HDMI', trackingMode: 'Bulk', quantity: 20, quantityOut: 5 };

const laptopOut = {
  id: 11, assetId: 1, assetKey: laptop.assetKey, personEmail: 'amir@pmw.com',
  quantity: 1, returnedQuantity: 0, handoverStatus: HANDOVER_STATUS.OUT,
  kind: HANDOVER_KIND.ISSUED,
};

const cablesOut = {
  id: 12, assetId: 2, assetKey: cables.assetKey, personEmail: 'amir@pmw.com',
  quantity: 3, returnedQuantity: 0, handoverStatus: HANDOVER_STATUS.OUT,
  kind: HANDOVER_KIND.ISSUED,
};

const handovers = [laptopOut, cablesOut];
const register = [laptop, cables];

describe('returning a tracked item', () => {
  it('closes the handover and puts the item back on the shelf', () => {
    const plan = planReturn(
      [{ handoverId: 11, quantity: 1, condition: 'Good' }],
      handovers, register, { returnedOn: 99, returnedBy: 'it@pmw' },
    );

    expect(plan.handoverUpdates[0].body).toMatchObject({
      returnedQuantity: 1,
      handoverStatus: HANDOVER_STATUS.RETURNED,
      returnedOn: 99,
      returnedBy: 'it@pmw',
      returnCondition: 'Good',
    });
    expect(plan.assetUpdates[0].body).toMatchObject({
      quantityOut: 0, status: 'In stock', assignedTo: '', condition: 'Good',
    });
  });

  /** A monitor returned faulty must not rejoin the shelf looking available. */
  it('records the condition it came back in', () => {
    const plan = planReturn([{ handoverId: 11, condition: 'Faulty' }], handovers, register);
    expect(plan.assetUpdates[0].body.condition).toBe('Faulty');
  });

  it('returns everything still out when no quantity is named', () => {
    const plan = planReturn([{ handoverId: 11 }], handovers, register);
    expect(plan.handoverUpdates[0].body.returnedQuantity).toBe(1);
  });
});

describe('returning bulk stock', () => {
  it('brings down what is out without touching what is owned', () => {
    const plan = planReturn([{ handoverId: 12, quantity: 3 }], handovers, register);

    expect(plan.assetUpdates[0].body.quantityOut).toBe(2);
    expect('quantity' in plan.assetUpdates[0].body).toBe(false);
  });

  it('leaves a partly returned line partly returned', () => {
    const plan = planReturn([{ handoverId: 12, quantity: 1 }], handovers, register);

    expect(plan.handoverUpdates[0].body.handoverStatus).toBe(HANDOVER_STATUS.PARTLY);
    expect(plan.assetUpdates[0].body.quantityOut).toBe(4);
  });

  it('does not call a box back in stock while some of it is still out', () => {
    const plan = planReturn([{ handoverId: 12, quantity: 3 }], handovers, register);
    expect(plan.assetUpdates[0].body.status).toBeUndefined();
  });

  it('calls it in stock once the last one is back', () => {
    const emptied = [{ ...cables, quantityOut: 3 }];
    const plan = planReturn([{ handoverId: 12, quantity: 3 }], handovers, emptied);

    expect(plan.assetUpdates[0].body.status).toBe('In stock');
  });

  /**
   * Two lines of the same box coming back together have to accumulate against
   * one register row, or the second write undoes the first's arithmetic.
   */
  it('accumulates two lines of the same box into one register write', () => {
    const second = { ...cablesOut, id: 13, quantity: 2 };
    const plan = planReturn(
      [{ handoverId: 12, quantity: 3 }, { handoverId: 13, quantity: 2 }],
      [...handovers, second], register,
    );

    expect(plan.assetUpdates).toHaveLength(1);
    expect(plan.assetUpdates[0].body.quantityOut).toBe(0);
    expect(plan.handoverUpdates).toHaveLength(2);
  });
});

describe('what a return refuses', () => {
  /** "Two came back" and "three came back" is a real disagreement, not a rounding. */
  it('refuses more than is still out, rather than clamping it', () => {
    const plan = planReturn([{ handoverId: 12, quantity: 9 }], handovers, register);

    expect(plan.handoverUpdates).toHaveLength(0);
    expect(plan.blocked[0].reason).toBe('Only 3 still out on that handover.');
  });

  it('refuses a line already fully returned', () => {
    const done = [{ ...cablesOut, returnedQuantity: 3 }];
    const plan = planReturn([{ handoverId: 12 }], done, register);

    expect(plan.blocked[0].reason).toContain('Nothing to return');
  });

  it('refuses a handover that is no longer there', () => {
    const plan = planReturn([{ handoverId: 999 }], handovers, register);
    expect(plan.blocked[0].reason).toContain('no longer in the list');
  });

  it('still records the handover when the register row has vanished', () => {
    const plan = planReturn([{ handoverId: 11 }], handovers, []);

    expect(plan.handoverUpdates).toHaveLength(1);
    expect(plan.assetUpdates).toHaveLength(0);
  });

  it('lets the good lines through', () => {
    const plan = planReturn(
      [{ handoverId: 12, quantity: 9 }, { handoverId: 11 }],
      handovers, register,
    );

    expect(plan.blocked).toHaveLength(1);
    expect(plan.handoverUpdates).toHaveLength(1);
  });
});

describe('returnEverything', () => {
  it('lists every open line with what is still out on it', () => {
    const entries = returnEverything([laptopOut, { ...cablesOut, returnedQuantity: 1 }]);

    expect(entries).toEqual([
      { handoverId: 11, quantity: 1, condition: '' },
      { handoverId: 12, quantity: 2, condition: '' },
    ]);
  });

  it('leaves out what is already back', () => {
    const entries = returnEverything([{ ...laptopOut, returnedQuantity: 1 }]);
    expect(entries).toEqual([]);
  });

  it('carries a condition onto every line', () => {
    expect(returnEverything([laptopOut], 'Good')[0].condition).toBe('Good');
  });
});

describe('signing for a return', () => {
  it('records where the signature is when there is one', () => {
    const plan = planReturn(
      [{ handoverId: 11, quantity: 1 }],
      handovers, register, { returnSignature: '/sites/it/Photos/signature-1.png' },
    );

    expect(plan.handoverUpdates[0].body.returnSignature)
      .toBe('/sites/it/Photos/signature-1.png');
  });

  it('leaves the field alone entirely when nobody signed', () => {
    const plan = planReturn([{ handoverId: 11, quantity: 1 }], handovers, register);

    // Not '' -- a blank would rub out the signature on the half returned first.
    expect('returnSignature' in plan.handoverUpdates[0].body).toBe(false);
  });
});
