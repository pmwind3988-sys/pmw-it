import { describe, it, expect } from 'vitest';
import { planCombine, blockersFor, differencesIn } from './combine.js';
import { parseUnits } from './units.js';

const tracked = (id, serial, extra = {}) => ({
  id,
  trackingMode: 'Tracked',
  category: 'Monitor',
  manufacturer: 'Dell',
  model: 'P2422H',
  serialNumber: serial,
  quantity: 1,
  arrivedOn: 1000 + id,
  assetKey: `serial:dell|${serial}`,
  ...extra,
});

const ten = () => Array.from({ length: 10 }, (unused, at) => tracked(at + 1, `SN-${at + 1}`));

describe('planCombine', () => {
  it('turns ten rows of one monitor into one line of ten', () => {
    const plan = planCombine(ten());

    expect(plan.edits.quantity).toBe(10);
    expect(plan.edits.trackingMode).toBe('Bulk');
    expect(plan.remove).toHaveLength(9);
  });

  it('keeps every serial number, one per item', () => {
    const plan = planCombine(ten());
    const units = parseUnits(plan.edits.units);

    expect(units).toHaveLength(10);
    expect(units.map((unit) => unit.serialNumber).sort())
      .toEqual(ten().map((row) => row.serialNumber).sort());
  });

  it('folds into the oldest row, so the longest history survives', () => {
    const rows = [tracked(7, 'SN-7'), tracked(2, 'SN-2'), tracked(9, 'SN-9')];

    expect(planCombine(rows).keep.id).toBe(2);
    expect(planCombine(rows).remove.map((row) => row.id)).toEqual([7, 9]);
  });

  it('gives each row its own block of items rather than overwriting item 1', () => {
    const bulk = {
      id: 1,
      trackingMode: 'Bulk',
      category: 'Monitor',
      quantity: 3,
      arrivedOn: 10,
      units: JSON.stringify([{ index: 0, serialNumber: 'A' }, { index: 2, serialNumber: 'C' }]),
    };
    const plan = planCombine([bulk, tracked(2, 'SN-late', { arrivedOn: 99 })]);

    expect(plan.edits.quantity).toBe(4);
    const units = parseUnits(plan.edits.units);
    expect(units.map((unit) => unit.serialNumber)).toEqual(['A', 'C', 'SN-late']);
    // The late row's serial takes a position of its own past the line of three.
    expect(units.find((unit) => unit.serialNumber === 'SN-late').index).toBe(3);
  });

  it('carries a row\'s condition and status onto its own item', () => {
    const rows = [
      tracked(1, 'SN-1', { condition: 'New', status: 'In store' }),
      tracked(2, 'SN-2', { condition: 'Faulty', status: 'Under repair' }),
    ];
    const units = parseUnits(planCombine(rows).edits.units);

    expect(units.find((unit) => unit.serialNumber === 'SN-2').condition).toBe('Faulty');
    expect(units.find((unit) => unit.serialNumber === 'SN-2').status).toBe('Under repair');
  });

  it('says so when the rows are not obviously the same thing', () => {
    const rows = [tracked(1, 'SN-1'), tracked(2, 'SN-2', { model: 'U2723QE' })];

    expect(planCombine(rows).warnings).toEqual(['model']);
    expect(differencesIn([tracked(1, 'A'), tracked(2, 'B')])).toEqual([]);
  });
});

describe('blockersFor', () => {
  const rows = [tracked(1, 'SN-1'), tracked(2, 'SN-2')];

  it('allows a straightforward pair', () => {
    expect(blockersFor(rows, [])).toEqual([]);
  });

  it('needs two rows to have anything to combine', () => {
    expect(blockersFor([rows[0]], [])).toHaveLength(1);
  });

  it('refuses while one of them is out with somebody', () => {
    const handovers = [{ assetKey: rows[1].assetKey, quantity: 1, returnedQuantity: 0 }];

    expect(blockersFor(rows, handovers)).toHaveLength(1);
    expect(blockersFor(rows, handovers)[0]).toMatch(/out with somebody/);
  });

  it('is happy once it has been brought back', () => {
    const handovers = [{ assetKey: rows[1].assetKey, quantity: 1, returnedQuantity: 1 }];

    expect(blockersFor(rows, handovers)).toEqual([]);
  });
});
