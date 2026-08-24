import { describe, it, expect } from 'vitest';
import { planSave, coalesce, diffAsset, applyManualOverrides } from './planSave.js';
import { newDraft } from '../draft/draftAsset.js';
import { TRACKED, BULK } from '../assetKinds.js';

const laptop = (overrides = {}) => newDraft({
  category: 'Laptop',
  manufacturer: 'Dell',
  model: 'Latitude 5540',
  serialNumber: 'CN0ABC123',
  ...overrides,
});

const mice = (overrides = {}) => newDraft({
  category: 'Mouse',
  manufacturer: 'Logitech',
  model: 'B100',
  quantity: 10,
  ...overrides,
});

/** Two tabs already in the register, as a bulk line with a unit record. */
const tabs = (overrides = {}) => ({
  id: 3,
  assetKey: 'bulk:TAB|SAMSUNG|SM-X210',
  trackingMode: BULK,
  category: 'Tab',
  manufacturer: 'Samsung',
  model: 'SM-X210',
  quantity: 2,
  ...overrides,
});

describe('planSave — new things', () => {
  it('inserts a row nothing in the register matches', () => {
    const plan = planSave([laptop()], []);

    expect(plan.inserts).toHaveLength(1);
    expect(plan.updates).toHaveLength(0);
    expect(plan.inserts[0].assetKey).toBe('serial:DELL|CN0ABC123');
    expect(plan.inserts[0].body.title).toBe('Dell Latitude 5540 — CN0ABC123');
  });

  it('stamps who added it and when', () => {
    const plan = planSave([laptop()], [], { addedOn: 42, addedBy: 'ashraf@pmw' });

    expect(plan.inserts[0].body.addedOn).toBe(42);
    expect(plan.inserts[0].body.addedBy).toBe('ashraf@pmw');
  });
});

describe('planSave — things already in the register', () => {
  /** The dedupe that matters: the same machine scanned again is one row. */
  it('updates a tracked row rather than making a second one', () => {
    const register = [{
      id: 7, assetKey: 'serial:DELL|CN0ABC123', trackingMode: TRACKED,
      category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5540', quantity: 1,
      location: 'Store room',
    }];
    const plan = planSave([laptop({ location: 'Level 3' })], register);

    expect(plan.inserts).toHaveLength(0);
    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].id).toBe(7);
    expect(plan.updates[0].body.location).toBe('Level 3');
  });

  /** A second bag of the same mice is more stock, not a correction. */
  it('adds to the quantity of a bulk line', () => {
    const register = [{
      id: 3, assetKey: 'bulk:MOUSE|LOGITECH|B100', trackingMode: BULK,
      category: 'Mouse', manufacturer: 'Logitech', model: 'B100', quantity: 10,
    }];
    const plan = planSave([mice({ quantity: 5 })], register);

    expect(plan.updates[0].body.quantity).toBe(15);
  });

  /**
   * A save writes every column, and a draft has no unit records — nothing in a
   * barcode says which of the two tabs is which. Without carrying them across,
   * a second box of the same thing would erase every serial anybody had
   * recorded against the individual items on that row.
   */
  it('carries the unit records of a bulk row through a re-scan', () => {
    const units = JSON.stringify([{ index: 0, serialNumber: 'R52TC0AAAAA' }]);
    const plan = planSave([newDraft({
      category: 'Tab', manufacturer: 'Samsung', model: 'SM-X210', quantity: 1,
    })], [tabs({ units })]);

    expect(plan.updates[0].body.quantity).toBe(3);
    expect(plan.updates[0].body.units).toBe(units);
  });

  /**
   * A third tab arriving is a THIRD object. Its serial takes the next position
   * rather than filling a gap in item 1's record, which would invent an item
   * wearing one tab's serial and another's label.
   */
  it('gives a re-scanned bulk item the next position, not the first', () => {
    const units = JSON.stringify([{ index: 0, serialNumber: 'R52TC0AAAAA' }]);
    const plan = planSave([newDraft({
      category: 'Tab', manufacturer: 'Samsung', model: 'SM-X210', quantity: 1,
      serialNumber: 'R52TC0CCCCC',
    })], [tabs({ units })]);

    expect(JSON.parse(plan.updates[0].body.units)).toEqual([
      expect.objectContaining({ index: 0, serialNumber: 'R52TC0AAAAA' }),
      expect.objectContaining({ index: 2, serialNumber: 'R52TC0CCCCC' }),
    ]);
    // And never on the row: one serial cannot speak for three tabs.
    expect(plan.updates[0].body.serialNumber).toBe('');
  });

  it('leaves a tracked row at one however it was scanned', () => {
    const register = [{
      id: 7, assetKey: 'serial:DELL|CN0ABC123', trackingMode: TRACKED, quantity: 1,
      category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5540',
    }];
    const plan = planSave([laptop({ quantity: 4 })], register);

    expect(plan.updates[0]?.body.quantity ?? 1).toBe(1);
  });

  it('writes nothing at all when nothing changed', () => {
    const register = [{
      id: 7, assetKey: 'serial:DELL|CN0ABC123', trackingMode: TRACKED, quantity: 1,
      category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5540',
      serialNumber: 'CN0ABC123', condition: 'New',
    }];
    const plan = planSave([laptop()], register);

    expect(plan.updates).toHaveLength(0);
    expect(plan.unchanged).toBe(1);
  });

  it('records what changed, field by field', () => {
    const register = [{
      id: 7, assetKey: 'serial:DELL|CN0ABC123', trackingMode: TRACKED, quantity: 1,
      category: 'Laptop', manufacturer: 'Dell', model: 'Latitude 5530',
      serialNumber: 'CN0ABC123', condition: 'New',
    }];
    const plan = planSave([laptop()], register);
    const change = plan.changeRows.find((row) => row.fieldName === 'model');

    expect(change).toMatchObject({
      oldValue: 'Latitude 5530', newValue: 'Latitude 5540', changeType: 'Updated',
    });
  });
});

describe('planSave — sticker labels', () => {
  it('blocks a row whose label is already on something else, and saves the rest', () => {
    const register = [{
      id: 9, assetKey: 'serial:HP|OTHER', title: 'HP EliteBook', assetTag: 'PMW-0142',
    }];
    const plan = planSave([laptop({ assetTag: 'pmw-0142' }), mice()], register);

    expect(plan.blocked).toHaveLength(1);
    expect(plan.blocked[0].issues[0].message).toContain('HP EliteBook');
    expect(plan.inserts).toHaveLength(1);
    expect(plan.inserts[0].body.category).toBe('Mouse');
  });

  it('blocks the second of two rows in one batch claiming the same label', () => {
    const plan = planSave([
      laptop({ assetTag: 'PMW-0142' }),
      laptop({ serialNumber: 'CN0XYZ999', assetTag: 'PMW-0142' }),
    ], []);

    expect(plan.inserts).toHaveLength(1);
    expect(plan.blocked).toHaveLength(1);
  });

  it('lets a row keep the label it already has in the register', () => {
    const register = [{
      id: 7, assetKey: 'serial:DELL|CN0ABC123', trackingMode: TRACKED,
      assetTag: 'PMW-0142', quantity: 1, category: 'Laptop',
    }];
    // The clash check is against OTHER rows; a row keeping its own label is
    // the ordinary case of re-scanning a labelled machine.
    const plan = planSave([laptop({ assetTag: 'PMW-0142' })], register);

    expect(plan.blocked).toHaveLength(0);
  });
});

describe('coalesce', () => {
  it('adds up two bulk rows of the same model', () => {
    const [row] = coalesce([mice({ quantity: 10 }), mice({ quantity: 5 })]);
    expect(row.quantity).toBe(15);
  });

  it('folds two scans of the same tracked unit into one row', () => {
    const rows = coalesce([laptop(), laptop({ model: '' , partNumber: '5UF44AA' })]);

    expect(rows).toHaveLength(1);
    expect(rows[0].partNumber).toBe('5UF44AA');
    expect(rows[0].model).toBe('Latitude 5540');
  });

  it('keeps genuinely different things apart', () => {
    expect(coalesce([laptop(), laptop({ serialNumber: 'CN0XYZ999' })])).toHaveLength(2);
  });

  /** Two unserialised, unlabelled units are two units, not one. */
  it('does not fold rows that have no stable identity', () => {
    const rows = coalesce([
      newDraft({ category: 'Laptop', model: 'Spare' }),
      newDraft({ category: 'Laptop', model: 'Spare' }),
    ]);

    expect(rows).toHaveLength(2);
  });
});

describe('diffAsset', () => {
  it('calls a first value Added and a cleared one Removed', () => {
    expect(diffAsset({ location: '' }, { location: 'Store room' })[0].changeType).toBe('Added');
    expect(diffAsset({ location: 'Store room' }, { location: '' })[0].changeType).toBe('Removed');
  });

  it('does not report a field that only differs by being null rather than blank', () => {
    expect(diffAsset({ location: null }, { location: '' })).toEqual([]);
  });

  it('ignores fields nobody asked it to track', () => {
    expect(diffAsset({ photoUrl: 'a' }, { photoUrl: 'b' })).toEqual([]);
  });
});

describe('applyManualOverrides', () => {
  /** Otherwise the edit screen is a trap: the next scan undoes the correction. */
  it('holds back a field somebody corrected by hand', () => {
    const merged = applyManualOverrides(
      { model: 'from the barcode', location: 'Level 3' },
      { model: 'typed by a person', manualFields: ['model'] },
    );

    expect(merged.model).toBe('typed by a person');
    expect(merged.location).toBe('Level 3');
  });

  it('carries the list forward so updating something else does not wipe it', () => {
    const merged = applyManualOverrides({ model: 'x' }, { model: 'y', manualFields: ['model'] });
    expect(merged.manualFields).toEqual(['model']);
  });

  it('leaves an incoming record alone when nothing was hand-set', () => {
    const incoming = { model: 'x' };
    expect(applyManualOverrides(incoming, { manualFields: [] })).toBe(incoming);
  });
});
