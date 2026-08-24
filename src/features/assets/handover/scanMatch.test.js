import { describe, it, expect } from 'vitest';
import { findScanTarget } from './scanMatch.js';
import { serialiseUnits } from '../units.js';
import { TRACKED, BULK } from '../assetKinds.js';

const bulkTabs = {
  id: 1,
  assetKey: 'tabs',
  title: 'Lenovo Tab',
  trackingMode: BULK,
  quantity: 2,
  units: serialiseUnits([
    { index: 0, serialNumber: 'TAB-AAA' },
    { index: 1, serialNumber: 'TAB-BBB' },
  ]),
};

const laptop = {
  id: 2,
  assetKey: 'laptop',
  title: 'ThinkPad',
  trackingMode: TRACKED,
  serialNumber: 'LT-999',
};

describe('findScanTarget', () => {
  it('resolves a bulk code to the exact unit it names', () => {
    const target = findScanTarget([bulkTabs, laptop], 'TAB-BBB');
    expect(target.asset.id).toBe(1);
    expect(target.unit.index).toBe(1);
  });

  it('tells two units of the same box apart', () => {
    expect(findScanTarget([bulkTabs], 'tab-aaa').unit.index).toBe(0);
    expect(findScanTarget([bulkTabs], 'tab-bbb').unit.index).toBe(1);
  });

  it('reads a tracked serial as the whole row, never as a unit', () => {
    const target = findScanTarget([bulkTabs, laptop], 'LT-999');
    expect(target.asset.id).toBe(2);
    expect(target.unit).toBeNull();
  });

  it('answers with nothing for a code the register does not carry', () => {
    expect(findScanTarget([bulkTabs, laptop], 'NOPE')).toBeNull();
  });

  it('reads a bulk row-level tag as unit one, since that is where it belongs', () => {
    // A tag sitting on a bulk row is a leftover that describes item 1; the
    // register surfaces it as unit 0, and scanning it hands out that one item.
    const boxed = {
      id: 3, assetKey: 'cables', title: 'HDMI cable', trackingMode: BULK, quantity: 10, assetTag: 'BOX-1',
    };
    const target = findScanTarget([boxed], 'BOX-1');
    expect(target.asset.id).toBe(3);
    expect(target.unit.index).toBe(0);
  });

  it('matches a bulk row as a whole on a model-wide extra code, not as a unit', () => {
    const boxed = {
      id: 4, assetKey: 'mice', title: 'USB mouse', trackingMode: BULK, quantity: 10, additionalCodes: ['MODEL-M100'],
    };
    const target = findScanTarget([boxed], 'MODEL-M100');
    expect(target.asset.id).toBe(4);
    expect(target.unit).toBeNull();
  });
});
