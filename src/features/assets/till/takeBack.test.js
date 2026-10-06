import { describe, it, expect } from 'vitest';
import { TRACKED, BULK } from '../assetKinds.js';
import {
  scanBack, BACK_RESULT, chooseHolder, addHandover, setReturnQuantity, toReturns, itemCount,
} from './takeBack.js';

const LAPTOP = { id: 1, assetKey: 'serial:DELL|A1', title: 'Dell Latitude', trackingMode: TRACKED, serialNumber: 'A1', quantity: 1 };
const MICE = { id: 2, assetKey: 'bulk:Mouse||M90', title: 'M90', trackingMode: BULK, partNumber: '509920', quantity: 10 };
const SPARE = { id: 3, assetKey: 'serial:DELL|C3', title: 'Dell Latitude', trackingMode: TRACKED, serialNumber: 'C3', quantity: 1 };
const REGISTER = [LAPTOP, MICE, SPARE];

const H = (over) => ({ quantity: 1, returnedQuantity: 0, unitIndex: null, serialNumber: '', ...over });
const HANDOVERS = [
  H({ id: 10, assetKey: LAPTOP.assetKey, personName: 'Daniel Lim', personEmail: 'd@x' }),
  H({ id: 11, assetKey: MICE.assetKey, personName: 'Nurul Huda', personEmail: 'n@x', quantity: 2 }),
  H({ id: 12, assetKey: MICE.assetKey, personName: 'Hafiz', personEmail: 'h@x', quantity: 1 }),
  H({ id: 13, assetKey: LAPTOP.assetKey, personName: 'Old', returnedQuantity: 1 }),
];

describe('scanBack', () => {
  it('finds who has a laptop without being told', () => {
    const { lines, result, line } = scanBack([], 'a1', REGISTER, HANDOVERS);
    expect(result).toBe(BACK_RESULT.ADDED);
    expect(line).toMatchObject({ handoverId: 10, from: 'Daniel Lim', quantity: 1, single: true });
    expect(scanBack(lines, 'A1', REGISTER, HANDOVERS).result).toBe(BACK_RESULT.DUPLICATE);
  });

  it('refuses something that is not out with anybody', () => {
    expect(scanBack([], 'C3', REGISTER, HANDOVERS).result).toBe(BACK_RESULT.NOT_OUT);
    expect(scanBack([], 'nope', REGISTER, HANDOVERS).result).toBe(BACK_RESULT.MISSING);
  });

  it('asks whose it is when several people hold the line, then counts up to what they have', () => {
    let { lines, result, line } = scanBack([], '509920', REGISTER, HANDOVERS);
    expect(result).toBe(BACK_RESULT.CHOOSE);
    expect(line.choices.map((choice) => choice.handoverId)).toEqual([11, 12]);
    expect(toReturns(lines, 'Good')).toEqual([]);

    lines = chooseHolder(lines, line.lineId, 11, HANDOVERS, REGISTER);
    expect(lines[0]).toMatchObject({ handoverId: 11, from: 'Nurul Huda', max: 2 });

    ({ lines, result } = scanBack(lines, '509920', REGISTER, HANDOVERS));
    expect(result).toBe(BACK_RESULT.COUNTED);
    ({ result } = scanBack(lines, '509920', REGISTER, HANDOVERS));
    expect(result).toBe(BACK_RESULT.ALL_BACK);
    expect(toReturns(lines, 'Fair')).toEqual([{ handoverId: 11, quantity: 2, condition: 'Fair' }]);
  });

  it('takes a whole handover from a list and keeps the count inside what is out', () => {
    let { lines } = addHandover([], HANDOVERS[1], REGISTER);
    expect(lines[0].quantity).toBe(2);
    lines = setReturnQuantity(lines, lines[0].lineId, 9);
    expect(lines[0].quantity).toBe(2);
    expect(itemCount(lines)).toBe(2);
  });
});
