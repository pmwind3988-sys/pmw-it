import { describe, it, expect } from 'vitest';
import { TRACKED, BULK } from '../assetKinds.js';
import { tillSearch, quickKeys } from './tillSearch.js';

const A = (over) => ({ quantityOut: 0, quantity: 1, ...over });
const ASSETS = [
  A({ id: 1, assetKey: 'k1', title: 'Dell Latitude 5450', category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5450', trackingMode: TRACKED, serialNumber: '5CG1', assetTag: 'PMW-0142' }),
  A({ id: 2, assetKey: 'k2', title: 'Dell Latitude 5450', category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5450', trackingMode: TRACKED, serialNumber: '5CG2', quantityOut: 1 }),
  A({ id: 3, assetKey: 'k3', title: 'Logitech M90', category: 'Mouse', manufacturer: 'Logitech', model: 'M90', trackingMode: BULK, quantity: 17, quantityOut: 2 }),
  A({ id: 4, assetKey: 'k4', title: 'HDMI cable', category: 'Cable', model: 'HDMI 1.8 m', trackingMode: BULK, quantity: 40 }),
];
const HANDOVERS = [
  { id: 9, assetKey: 'k2', itemTitle: 'Dell Latitude 5450', serialNumber: '5CG2', personName: 'Daniel Lim', quantity: 1, returnedQuantity: 0 },
  { id: 8, assetKey: 'k3', itemTitle: 'Logitech M90', personName: 'Nurul Huda', quantity: 2, returnedQuantity: 0 },
  { id: 7, assetKey: 'k3', itemTitle: 'Logitech M90', personName: 'Old', quantity: 1, returnedQuantity: 1 },
];

describe('tillSearch', () => {
  it('waits for two characters', () => {
    expect(tillSearch('out', 'd', { assets: ASSETS })).toEqual([]);
  });

  it('offers each known model once for a delivery', () => {
    const found = tillSearch('in', '5450', { assets: ASSETS });
    expect(found).toHaveLength(1);
    expect(found[0]).toMatchObject({ kind: 'model', name: 'Dell Latitude 5450' });
  });

  it('offers only what is left to hand out, found by label too', () => {
    expect(tillSearch('out', 'latitude', { assets: ASSETS }).map((r) => r.asset.id)).toEqual([1]);
    expect(tillSearch('out', '0142', { assets: ASSETS }).map((r) => r.asset.id)).toEqual([1]);
  });

  it('lists each item whose serial matches part of what was typed, with the full serial', () => {
    const tabs = A({
      id: 5, assetKey: 'k5', title: 'Galaxy Tab A9', category: 'Tablet', manufacturer: 'Samsung', model: 'Tab A9',
      trackingMode: BULK, quantity: 4, quantityOut: 1,
      units: JSON.stringify([
        { index: 0, serialNumber: 'R9XW1001' },
        { index: 1, serialNumber: 'R9XW1002' },
        { index: 2, serialNumber: 'R9XW2003' },
        { index: 3, assetTag: 'TAB-1009' },
      ]),
    });
    const handovers = [{ id: 1, assetKey: 'k5', unitIndex: 1, quantity: 1, returnedQuantity: 0 }];
    const found = tillSearch('out', 'xw10', { assets: [tabs], handovers });
    expect(found.map((r) => [r.kind, r.unit.index])).toEqual([['unit', 0]]);
    expect(found[0].sub).toContain('R9XW1001');
    expect(found[0].name).toContain('Galaxy Tab A9');

    // A unit's sticker label is found the same way; the product name still finds the line.
    expect(tillSearch('out', '1009', { assets: [tabs], handovers }).map((r) => r.kind)).toEqual(['unit']);
    expect(tillSearch('out', 'galaxy', { assets: [tabs], handovers }).map((r) => r.kind)).toEqual(['asset']);
  });

  it('finds a return by the thing or by who has it, open handovers only', () => {
    expect(tillSearch('back', 'daniel', { assets: ASSETS, handovers: HANDOVERS }).map((r) => r.handover.id)).toEqual([9]);
    expect(tillSearch('back', 'm90', { assets: ASSETS, handovers: HANDOVERS }).map((r) => r.handover.id)).toEqual([8]);
  });
});

describe('quickKeys', () => {
  it('ranks bulk lines by how often they go out, then by stock', () => {
    const keys = quickKeys(ASSETS, HANDOVERS);
    expect(keys.map((key) => key.asset.id)).toEqual([3, 4]);
    expect(keys[0].left).toBe(15);
  });
});
