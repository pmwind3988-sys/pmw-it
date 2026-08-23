import { describe, it, expect } from 'vitest';
import {
  owned, out, available, isOut, outstanding, isOpen, isOverdue, statusFor,
  holdersOf, heldBy, peopleWithItems, HANDOVER_KIND, HANDOVER_STATUS,
} from './availability.js';

const box = (overrides = {}) => ({
  trackingMode: 'Bulk', quantity: 20, quantityOut: 3, ...overrides,
});

const laptop = (overrides = {}) => ({
  trackingMode: 'Tracked', quantity: 1, quantityOut: 0, ...overrides,
});

const line = (overrides = {}) => ({
  assetKey: 'bulk:CABLE||HDMI',
  personEmail: 'amir@pmw.com',
  personName: 'Amir',
  quantity: 3,
  returnedQuantity: 0,
  kind: HANDOVER_KIND.ISSUED,
  ...overrides,
});

describe('what is on the shelf', () => {
  /** The figure the register exists to answer stays put when stock goes out. */
  it('keeps owned as what the company bought', () => {
    expect(owned(box())).toBe(20);
    expect(available(box())).toBe(17);
  });

  it('treats a tracked item as one, whatever its quantity says', () => {
    expect(owned(laptop({ quantity: 9 }))).toBe(1);
  });

  /** Rows saved before the column existed read as null. */
  it('reads a missing out-count as none out', () => {
    expect(out({ quantity: 5 })).toBe(0);
    expect(available({ trackingMode: 'Bulk', quantity: 5 })).toBe(5);
  });

  it('reads a missing quantity as one', () => {
    expect(owned({ trackingMode: 'Bulk' })).toBe(1);
  });

  /** A bad figure should read as "none left", never as a credit. */
  it('never reports a negative availability', () => {
    expect(available(box({ quantity: 2, quantityOut: 5 }))).toBe(0);
  });

  it('knows when something is out', () => {
    expect(isOut(box())).toBe(true);
    expect(isOut(laptop())).toBe(false);
    expect(isOut(laptop({ quantityOut: 1 }))).toBe(true);
  });
});

describe('one handover row', () => {
  it('counts what is still with the person', () => {
    expect(outstanding(line({ quantity: 3, returnedQuantity: 1 }))).toBe(2);
    expect(isOpen(line({ quantity: 3, returnedQuantity: 3 }))).toBe(false);
  });

  it('never reports a negative outstanding', () => {
    expect(outstanding(line({ quantity: 1, returnedQuantity: 4 }))).toBe(0);
  });

  it('moves through Out, Partly returned and Returned', () => {
    expect(statusFor(line({ quantity: 3, returnedQuantity: 0 }))).toBe(HANDOVER_STATUS.OUT);
    expect(statusFor(line({ quantity: 3, returnedQuantity: 1 }))).toBe(HANDOVER_STATUS.PARTLY);
    expect(statusFor(line({ quantity: 3, returnedQuantity: 3 }))).toBe(HANDOVER_STATUS.RETURNED);
  });
});

describe('overdue', () => {
  const borrowed = (dueOn) => line({ kind: HANDOVER_KIND.BORROWED, dueOn });

  it('is borrowed, still out, and past its date', () => {
    expect(isOverdue(borrowed(1000), 2000)).toBe(true);
    expect(isOverdue(borrowed(3000), 2000)).toBe(false);
  });

  /** The whole difference between the two kinds: Issued has no date to miss. */
  it('never applies to an issued item', () => {
    expect(isOverdue(line({ dueOn: 1000 }), 2000)).toBe(false);
  });

  it('does not chase something already back', () => {
    expect(isOverdue({ ...borrowed(1000), returnedQuantity: 3 }, 2000)).toBe(false);
  });

  it('does not chase a borrowed item with no date set', () => {
    expect(isOverdue(borrowed(null), 2000)).toBe(false);
  });
});

describe('who holds what', () => {
  const rows = [
    line({ assetKey: 'serial:DELL|A1', personEmail: 'amir@pmw.com', quantity: 1 }),
    line({ assetKey: 'serial:DELL|A1', personEmail: 'old@pmw.com', quantity: 1, returnedQuantity: 1 }),
    line({ personEmail: 'evonne@pmw.com', quantity: 2 }),
  ];

  it('names only the people who still have it', () => {
    const holders = holdersOf(rows, 'serial:DELL|A1');
    expect(holders).toHaveLength(1);
    expect(holders[0].personEmail).toBe('amir@pmw.com');
  });

  it('finds what one person holds, whatever case the email was typed in', () => {
    expect(heldBy(rows, 'AMIR@PMW.COM')).toHaveLength(1);
  });

  it('returns nothing for no email rather than everything', () => {
    expect(heldBy(rows, '')).toEqual([]);
    expect(heldBy(rows, null)).toEqual([]);
  });
});

describe('peopleWithItems', () => {
  /** Three cables and a laptop is four things, not two rows. */
  it('counts units rather than lines', () => {
    const people = peopleWithItems([
      line({ personEmail: 'amir@pmw.com', quantity: 3 }),
      line({ personEmail: 'amir@pmw.com', quantity: 1, assetKey: 'serial:DELL|A1' }),
    ]);

    expect(people).toHaveLength(1);
    expect(people[0].units).toBe(4);
    expect(people[0].lines).toBe(2);
  });

  it('puts whoever is overdue first', () => {
    const people = peopleWithItems([
      line({ personEmail: 'b@pmw.com', personName: 'Bee', quantity: 9 }),
      line({
        personEmail: 'a@pmw.com',
        personName: 'Ay',
        kind: HANDOVER_KIND.BORROWED,
        dueOn: 1000,
      }),
    ], 2000);

    expect(people[0].email).toBe('a@pmw.com');
    expect(people[0].overdue).toBe(1);
  });

  it('leaves out everything already returned', () => {
    expect(peopleWithItems([line({ quantity: 2, returnedQuantity: 2 })])).toEqual([]);
  });
});
