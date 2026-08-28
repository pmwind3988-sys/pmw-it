import { describe, it, expect } from 'vitest';
import { planHandover, coalesceLines, lineRefusal } from './planHandover.js';
import {
  newBasket, newLine, newUnitLine, addLine, resolveLines, setQuantity, hasAsset, unitCount,
} from './basket.js';
import { HANDOVER_KIND, HANDOVER_STATUS } from './availability.js';

const person = { name: 'Amir', email: 'amir@pmw.com', login: 'i:0#.f|membership|amir@pmw.com' };

const laptop = (overrides = {}) => ({
  id: 1,
  assetKey: 'serial:DELL|CN0ABC',
  title: 'Dell Latitude 5540 — CN0ABC',
  category: 'Laptop',
  trackingMode: 'Tracked',
  quantity: 1,
  quantityOut: 0,
  ...overrides,
});

const cables = (overrides = {}) => ({
  id: 2,
  assetKey: 'bulk:CABLE||HDMI',
  title: 'HDMI cable',
  category: 'Cable',
  trackingMode: 'Bulk',
  quantity: 20,
  quantityOut: 3,
  ...overrides,
});

const basketFor = (assets, overrides = {}) => assets.reduce(
  (basket, asset) => addLine(basket, newLine(asset)),
  { ...newBasket(person), ...overrides },
);

describe('the basket', () => {
  it('pins a tracked line to one however it is asked', () => {
    const basket = setQuantity(basketFor([laptop()]), null, 5);
    expect(basket.lines[0].quantity).toBe(1);
  });

  it('lets a bulk line take a quantity', () => {
    let basket = basketFor([cables()]);
    basket = setQuantity(basket, basket.lines[0].lineId, 4);
    expect(basket.lines[0].quantity).toBe(4);
    expect(unitCount(basket)).toBe(4);
  });

  it('refuses a quantity that is not a positive whole number', () => {
    let basket = basketFor([cables()]);
    basket = setQuantity(basket, basket.lines[0].lineId, '-2');
    expect(basket.lines[0].quantity).toBe(1);
  });

  it('knows the same laptop is already in it', () => {
    expect(hasAsset(basketFor([laptop()]), 1)).toBe(true);
    expect(hasAsset(basketFor([laptop()]), 99)).toBe(false);
  });

  it('gives every line the basket kind unless the line says otherwise', () => {
    const basket = { ...basketFor([laptop(), cables()]), kind: HANDOVER_KIND.BORROWED, dueOn: 500 };
    const lines = resolveLines(basket);

    expect(lines.every((line) => line.kind === HANDOVER_KIND.BORROWED)).toBe(true);
    expect(lines[0].dueOn).toBe(500);
  });

  /** An issued item has no date, so one left over from Borrowed is dropped. */
  it('drops a due date from a line that is issued rather than lent', () => {
    const basket = { ...basketFor([laptop()]), kind: HANDOVER_KIND.ISSUED, dueOn: 500 };
    expect(resolveLines(basket)[0].dueOn).toBeNull();
  });
});

describe('planHandover', () => {
  it('writes a handover row and moves the register row', () => {
    const plan = planHandover(basketFor([laptop()]), [laptop()], { issuedOn: 42, issuedBy: 'it@pmw' });

    expect(plan.handovers).toHaveLength(1);
    expect(plan.handovers[0]).toMatchObject({
      personEmail: 'amir@pmw.com',
      quantity: 1,
      handoverStatus: HANDOVER_STATUS.OUT,
      issuedOn: 42,
      issuedBy: 'it@pmw',
    });
    expect(plan.assetUpdates[0].body).toMatchObject({
      quantityOut: 1, status: 'Assigned', assignedTo: 'Amir',
    });
  });

  /** 20 owned, 3 already out, 4 more going out — still 20 owned. */
  it('adds to what is out without touching what is owned', () => {
    let basket = basketFor([cables()]);
    basket = setQuantity(basket, basket.lines[0].lineId, 4);
    const plan = planHandover(basket, [cables()]);

    expect(plan.assetUpdates[0].body.quantityOut).toBe(7);
    expect('quantity' in plan.assetUpdates[0].body).toBe(false);
  });

  /**
   * A bulk row can be with five people at once, so there is no honest single
   * value for who has it.
   */
  it('does not name a holder on a bulk row', () => {
    const plan = planHandover(basketFor([cables()]), [cables()]);
    expect(plan.assetUpdates[0].body.assignedTo).toBeUndefined();
  });

  it('marks a borrowed item borrowed and carries its due date', () => {
    const basket = { ...basketFor([laptop()]), kind: HANDOVER_KIND.BORROWED, dueOn: 900 };
    const plan = planHandover(basket, [laptop()]);

    expect(plan.handovers[0].dueOn).toBe(900);
    expect(plan.assetUpdates[0].body.status).toBe('Borrowed');
  });

  describe('refusals', () => {
    it('refuses a laptop somebody already has, and names them', () => {
      const held = laptop({ quantityOut: 1, assignedTo: 'Evonne' });
      const plan = planHandover(basketFor([held]), [held]);

      expect(plan.handovers).toHaveLength(0);
      expect(plan.blocked[0].reason).toContain('Evonne');
      expect(plan.blocked[0].conflictWith).toBe(1);
    });

    it('refuses more than is available, with the figure', () => {
      let basket = basketFor([cables({ quantityOut: 18 })]);
      basket = setQuantity(basket, basket.lines[0].lineId, 5);
      const plan = planHandover(basket, [cables({ quantityOut: 18 })]);

      expect(plan.blocked[0].reason).toBe('Only 2 of 20 available.');
    });

    /** One refusal is one line's problem. */
    it('hands over the rest of the basket anyway', () => {
      const held = laptop({ quantityOut: 1, assignedTo: 'Evonne' });
      const plan = planHandover(basketFor([held, cables()]), [held, cables()]);

      expect(plan.blocked).toHaveLength(1);
      expect(plan.handovers).toHaveLength(1);
      expect(plan.handovers[0].category).toBe('Cable');
    });

    it('refuses an item that has left the register since it was added', () => {
      const plan = planHandover(basketFor([laptop()]), []);
      expect(plan.blocked[0].reason).toContain('no longer in the register');
    });
  });
});

describe('coalesceLines', () => {
  /**
   * Two lines of three each pass individually against a stock of five and hand
   * out six. They have to be added up before the check, not after.
   */
  it('adds two lines of the same box together', () => {
    const [line] = coalesceLines([
      { assetId: 2, trackingMode: 'Bulk', quantity: 3, remarks: '' },
      { assetId: 2, trackingMode: 'Bulk', quantity: 3, remarks: '' },
    ]);

    expect(line.quantity).toBe(6);
  });

  it('catches the overdraw those two lines would have been', () => {
    const basket = {
      ...newBasket(person),
      lines: [
        { ...newLine(cables({ quantity: 5, quantityOut: 0 })), quantity: 3 },
        { ...newLine(cables({ quantity: 5, quantityOut: 0 })), quantity: 3 },
      ],
    };
    const plan = planHandover(basket, [cables({ quantity: 5, quantityOut: 0 })]);

    expect(plan.handovers).toHaveLength(0);
    expect(plan.blocked[0].reason).toBe('Only 5 of 5 available.');
  });

  it('lets the stricter kind win when two lines disagree', () => {
    const [line] = coalesceLines([
      { assetId: 2, trackingMode: 'Bulk', quantity: 1, kind: HANDOVER_KIND.ISSUED, remarks: '' },
      {
        assetId: 2, trackingMode: 'Bulk', quantity: 1, kind: HANDOVER_KIND.BORROWED,
        dueOn: 500, remarks: '',
      },
    ]);

    expect(line.kind).toBe(HANDOVER_KIND.BORROWED);
    expect(line.dueOn).toBe(500);
  });

  it('does not multiply a tracked item that somehow appears twice', () => {
    const [line] = coalesceLines([
      { assetId: 1, trackingMode: 'Tracked', quantity: 1, remarks: '' },
      { assetId: 1, trackingMode: 'Tracked', quantity: 1, remarks: '' },
    ]);

    expect(line.quantity).toBe(1);
  });
});

describe('lineRefusal', () => {
  it('says nothing when the line is fine', () => {
    const basket = basketFor([cables()]);
    expect(lineRefusal(basket.lines[0], cables(), basket)).toBeNull();
  });

  /** So the refusal appears while somebody is still standing at the desk. */
  it('counts what other lines of the basket have already claimed', () => {
    const stock = cables({ quantity: 5, quantityOut: 0 });
    const basket = {
      ...newBasket(person),
      lines: [
        { ...newLine(stock), lineId: 'a', quantity: 3 },
        { ...newLine(stock), lineId: 'b', quantity: 3 },
      ],
    };

    expect(lineRefusal(basket.lines[1], stock, basket)).toBe('Only 5 of 5 available.');
  });

  it('names whoever holds a tracked item', () => {
    const held = laptop({ quantityOut: 1, assignedTo: 'Evonne' });
    const basket = basketFor([held]);

    expect(lineRefusal(basket.lines[0], held, basket)).toContain('Evonne');
  });
});

describe('per-unit handover', () => {
  const tabs = (overrides = {}) => ({
    id: 3,
    assetKey: 'bulk:LENOVO||TAB',
    title: 'Lenovo Tab',
    category: 'Tablet',
    trackingMode: 'Bulk',
    quantity: 2,
    quantityOut: 0,
    ...overrides,
  });

  const unitLine = (asset, unit, overrides = {}) => ({
    ...newUnitLine(asset, unit),
    ...overrides,
  });

  it('writes one handover row per scanned unit, each with its own serial', () => {
    const asset = tabs();
    let basket = newBasket(person);
    basket = addLine(basket, unitLine(asset, { index: 0, serialNumber: 'TAB-AAA' }));
    basket = addLine(basket, unitLine(asset, { index: 1, serialNumber: 'TAB-BBB' }));

    const { handovers } = planHandover(basket, [asset]);

    expect(handovers).toHaveLength(2);
    expect(handovers.map((row) => row.serialNumber).sort()).toEqual(['TAB-AAA', 'TAB-BBB']);
    expect(handovers.every((row) => row.quantity === 1)).toBe(true);
    expect(handovers.map((row) => row.unitIndex).sort()).toEqual([0, 1]);
  });

  it('moves the register by the number of units, in a single row update', () => {
    const asset = tabs({ quantity: 5, quantityOut: 1 });
    let basket = newBasket(person);
    basket = addLine(basket, unitLine(asset, { index: 0, serialNumber: 'TAB-AAA' }));
    basket = addLine(basket, unitLine(asset, { index: 1, serialNumber: 'TAB-BBB' }));

    const { assetUpdates } = planHandover(basket, [asset]);

    expect(assetUpdates).toHaveLength(1);
    expect(assetUpdates[0].body.quantityOut).toBe(3);
  });

  it('refuses the unit that would overdraw the row, and keeps the ones that fit', () => {
    const asset = tabs({ quantity: 2, quantityOut: 1 });
    let basket = newBasket(person);
    basket = addLine(basket, unitLine(asset, { index: 0, serialNumber: 'TAB-AAA' }));
    basket = addLine(basket, unitLine(asset, { index: 1, serialNumber: 'TAB-BBB' }));

    const { handovers, blocked } = planHandover(basket, [asset]);

    expect(handovers).toHaveLength(1);
    expect(blocked).toHaveLength(1);
    expect(blocked[0].reason).toContain('Only 0 of 2 available');
  });

  it('does not name a holder on the bulk row when units go out', () => {
    const asset = tabs();
    let basket = newBasket(person);
    basket = addLine(basket, unitLine(asset, { index: 0, serialNumber: 'TAB-AAA' }));

    const { assetUpdates } = planHandover(basket, [asset]);

    expect(assetUpdates[0].body.assignedTo).toBeUndefined();
  });
});

describe('signing for a handover', () => {
  it('puts the signature on every line of the basket', () => {
    const basket = addLine(
      addLine(newBasket({ person }), newLine(laptop())),
      newLine(cables(), { quantity: 2 }),
    );
    const plan = planHandover(basket, [laptop(), cables()], {
      issueSignature: '/sites/it/Photos/signature-amir.png',
    });

    expect(plan.handovers).toHaveLength(2);
    for (const row of plan.handovers) {
      expect(row.issueSignature).toBe('/sites/it/Photos/signature-amir.png');
      // Nothing has come back yet, so nothing has been signed for coming back.
      expect(row.returnSignature).toBe('');
    }
  });

  it('records an unsigned handover as unsigned rather than refusing it', () => {
    const basket = addLine(newBasket({ person }), newLine(laptop()));
    const plan = planHandover(basket, [laptop()]);

    expect(plan.handovers).toHaveLength(1);
    expect(plan.handovers[0].issueSignature).toBe('');
  });
});
