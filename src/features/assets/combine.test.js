import { describe, it, expect } from 'vitest';
import { planCombine, blockersFor, differencesIn, stillOut } from './combine.js';
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
    expect(blockersFor(rows)).toEqual([]);
  });

  it('needs two rows to have anything to combine', () => {
    expect(blockersFor([rows[0]])).toHaveLength(1);
  });
});

describe('rows that are out with somebody', () => {
  const rows = [tracked(1, 'SN-1'), tracked(2, 'SN-2')];
  const held = {
    id: 55, assetId: 2, assetKey: rows[1].assetKey, itemTitle: 'Dell — SN-2',
    unitIndex: null, personEmail: 'aisyah@pmw', quantity: 1, returnedQuantity: 0,
  };

  it('combines them anyway, rather than sending somebody to fetch a monitor', () => {
    expect(blockersFor(rows)).toEqual([]);
    expect(planCombine(rows, [held]).edits.quantity).toBe(2);
  });

  it('says which ones are on a desk somewhere, so the screen can', () => {
    expect(stillOut(rows, [held])).toEqual([held]);
    expect(stillOut(rows, [{ ...held, returnedQuantity: 1 }])).toEqual([]);
  });

  it('points the handover at the item the row became', () => {
    const [move] = planCombine(rows, [held]).moves;

    expect(move.was).toBe(held);
    // SN-1 takes item 1, so the row that was out becomes item 2 — index 1.
    expect(move.unitIndex).toBe(1);
  });

  it('follows an item of a bulk row to its new number', () => {
    const first = tracked(1, 'SN-1');
    const box = {
      id: 2,
      trackingMode: 'Bulk',
      category: 'Monitor',
      manufacturer: 'Dell',
      model: 'P2422H',
      quantity: 2,
      arrivedOn: 9999,
      assetKey: 'bulk:MONITOR|DELL|P2422H',
      units: JSON.stringify([{ index: 1, serialNumber: 'B' }]),
    };
    const onUnitOne = { ...held, assetId: 2, assetKey: box.assetKey, unitIndex: 1 };

    const [move] = planCombine([first, box], [onUnitOne]).moves;

    // Item 2 of the box, behind SN-1's one item, is item 3 of the line.
    expect(move.unitIndex).toBe(2);
  });

  it('moves the returned ones too, so the history keeps its item', () => {
    const back = { ...held, id: 56, returnedQuantity: 1 };

    expect(planCombine(rows, [back]).moves).toHaveLength(1);
  });

  it('adds up what is out and stops the row naming one holder', () => {
    const out = [
      tracked(1, 'SN-1', { quantityOut: 1, assignedTo: 'Aisyah' }),
      tracked(2, 'SN-2', { quantityOut: 0 }),
      tracked(3, 'SN-3', { quantityOut: 1, assignedTo: 'Amir' }),
    ];
    const { edits } = planCombine(out);

    expect(edits.quantityOut).toBe(2);
    expect(edits.quantity).toBe(3);
    expect(edits.assignedTo).toBe('');
    expect(edits.assignedToEmail).toBe('');
  });

  it('tells two bulk rows of the same thing apart by row, not by key', () => {
    const box = (id, arrivedOn) => ({
      id,
      trackingMode: 'Bulk',
      category: 'Monitor',
      manufacturer: 'Dell',
      model: 'P2422H',
      quantity: 2,
      arrivedOn,
      assetKey: 'bulk:MONITOR|DELL|P2422H',
    });
    const onSecondRow = { ...held, assetId: 2, assetKey: 'bulk:MONITOR|DELL|P2422H', unitIndex: 0 };

    const [move] = planCombine([box(1, 10), box(2, 20)], [onSecondRow]).moves;

    // Item 1 of the SECOND row is item 3 of the line, not item 1 of it.
    expect(move.row.id).toBe(2);
    expect(move.unitIndex).toBe(2);
  });
});
