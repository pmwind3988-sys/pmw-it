import { describe, it, expect } from 'vitest';
import { newBasket } from '../handover/basket.js';
import { HANDOVER_KIND } from '../handover/availability.js';
import { TRACKED, BULK } from '../assetKinds.js';
import {
  scanOut, addAsset, OUT_RESULT, refusalsFor, sendable, itemCount, dueOnFor, withTerms,
} from './handOut.js';

const LAPTOP = {
  id: 1, assetKey: 'serial:DELL|A1', title: 'Dell Latitude', category: 'Laptop',
  trackingMode: TRACKED, serialNumber: 'A1', quantity: 1, quantityOut: 0,
};
const LENT = { ...LAPTOP, id: 3, assetKey: 'serial:DELL|B2', serialNumber: 'B2', quantityOut: 1 };
const MICE = {
  id: 2, assetKey: 'bulk:Mouse||M90', title: 'M90', category: 'Mouse', trackingMode: BULK,
  partNumber: '509920', quantity: 2, quantityOut: 0,
};
const REGISTER = [LAPTOP, LENT, MICE];

describe('scanOut', () => {
  it('adds a laptop once and refuses the same box again', () => {
    let { basket, result } = scanOut(newBasket(), 'a1', REGISTER);
    expect(result).toBe(OUT_RESULT.ADDED);
    ({ basket, result } = scanOut(basket, 'A1', REGISTER));
    expect(result).toBe(OUT_RESULT.DUPLICATE);
    expect(basket.lines).toHaveLength(1);
  });

  it('counts a bulk line up and flags it once it is more than there is', () => {
    let { basket } = scanOut(newBasket(), '509920', REGISTER);
    ({ basket } = scanOut(basket, '509920', REGISTER));
    expect(basket.lines[0].quantity).toBe(2);
    expect(refusalsFor(basket, REGISTER).size).toBe(0);

    ({ basket } = scanOut(basket, '509920', REGISTER));
    expect(refusalsFor(basket, REGISTER).size).toBe(1);
  });

  it('says why straight away, and leaves refused lines off what is sent', () => {
    let { basket } = scanOut(newBasket(), 'A1', REGISTER);
    ({ basket } = addAsset(basket, LENT));
    const refusals = refusalsFor(basket, REGISTER);
    expect(refusals.size).toBe(1);
    expect(sendable(basket, refusals).lines.map((line) => line.assetId)).toEqual([1]);
    expect(itemCount(basket, refusals)).toBe(1);
  });

  it('reports a code nothing in the register carries', () => {
    expect(scanOut(newBasket(), 'NOPE', REGISTER).result).toBe(OUT_RESULT.MISSING);
  });
});

describe('terms', () => {
  const now = new Date(2026, 9, 6, 9, 30).getTime();

  it('dates a loan at local noon and keeps no date for a keep', () => {
    const due = new Date(dueOnFor('2w', now));
    expect([due.getFullYear(), due.getMonth(), due.getDate(), due.getHours()]).toEqual([2026, 9, 20, 12]);
    expect(dueOnFor('none', now)).toBeNull();

    expect(withTerms(newBasket(), { loan: true, dueChoice: '1w', now }).kind).toBe(HANDOVER_KIND.BORROWED);
    expect(withTerms(newBasket(), { loan: false, dueChoice: '1w', now })).toMatchObject({ kind: HANDOVER_KIND.ISSUED, dueOn: null });
  });
});
