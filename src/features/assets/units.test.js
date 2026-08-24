import { describe, it, expect } from 'vitest';
import {
  parseUnits, unitsOf, setUnitField, serialiseUnits, diffUnits, filledCount, isBlankUnit,
  withUnitsSplitOut, mergeUnits, appendUnit, perItem, countPerItem, trimUnits, PER_UNIT_CODES,
} from './units.js';

const bulk = (quantity, units) => ({ quantity, trackingMode: 'Bulk', units });

describe('unitsOf', () => {
  it('gives one entry per unit the row says it owns', () => {
    expect(unitsOf(bulk(2))).toHaveLength(2);
    expect(unitsOf(bulk(5))).toHaveLength(5);
  });

  it('is one entry, not none, on a row with no quantity at all', () => {
    expect(unitsOf({})).toHaveLength(1);
  });

  it('fills the ones that were written on and leaves the rest blank', () => {
    const units = unitsOf(bulk(3, JSON.stringify([{ index: 1, serialNumber: 'SN-B' }])));

    expect(units[0].serialNumber).toBe('');
    expect(units[1].serialNumber).toBe('SN-B');
    expect(units[2].serialNumber).toBe('');
    expect(units.map((unit) => unit.index)).toEqual([0, 1, 2]);
  });

  /**
   * A quantity typed wrong and corrected back must not take a serial number
   * with it, so lowering the count only hides units.
   */
  it('hides units above the quantity without destroying them', () => {
    const stored = JSON.stringify([{ index: 2, serialNumber: 'SN-C' }]);

    expect(unitsOf(bulk(1, stored))).toHaveLength(1);
    expect(unitsOf(bulk(3, stored))[2].serialNumber).toBe('SN-C');
  });
});

describe('parseUnits', () => {
  it('reads nothing as nothing', () => {
    expect(parseUnits('')).toEqual([]);
    expect(parseUnits(null)).toEqual([]);
    expect(parseUnits(undefined)).toEqual([]);
  });

  /** A mangled column must still let the row open. */
  it('survives a value that is not JSON', () => {
    expect(parseUnits('{not json')).toEqual([]);
    expect(parseUnits('"a string"')).toEqual([]);
  });

  it('accepts an array it has already been handed', () => {
    expect(parseUnits([{ index: 0, assetTag: 'PMW-1' }])[0].assetTag).toBe('PMW-1');
  });

  it('drops entries with no usable position', () => {
    expect(parseUnits([{ serialNumber: 'X' }, { index: -1, serialNumber: 'Y' }])).toEqual([]);
  });

  it('drops entries nobody wrote anything on', () => {
    expect(parseUnits([{ index: 0, serialNumber: '   ' }])).toEqual([]);
  });
});

describe('serialiseUnits', () => {
  it('keeps only the units somebody filled in', () => {
    const units = unitsOf(bulk(4));
    const edited = setUnitField(units, 2, 'serialNumber', 'SN-C');

    const stored = JSON.parse(serialiseUnits(edited));
    expect(stored).toHaveLength(1);
    expect(stored[0]).toMatchObject({ index: 2, serialNumber: 'SN-C' });
  });

  it('is an empty string when nothing has been recorded', () => {
    expect(serialiseUnits(unitsOf(bulk(3)))).toBe('');
  });

  it('round-trips through the column unchanged', () => {
    const edited = setUnitField(unitsOf(bulk(2)), 1, 'assetTag', 'PMW-0142');
    expect(unitsOf(bulk(2, serialiseUnits(edited)))[1].assetTag).toBe('PMW-0142');
  });

  it('trims what it stores, so a stray space is not a value', () => {
    expect(serialiseUnits([{ index: 0, serialNumber: '  ' }])).toBe('');
  });
});

describe('setUnitField', () => {
  it('touches one unit and leaves its neighbours alone', () => {
    const units = unitsOf(bulk(3));
    const after = setUnitField(units, 1, 'condition', 'Faulty');

    expect(after[1].condition).toBe('Faulty');
    expect(after[0]).toBe(units[0]);
    expect(after[2]).toBe(units[2]);
  });
});

describe('diffUnits', () => {
  it('names the unit and the field, never the JSON', () => {
    const before = serialiseUnits(unitsOf(bulk(2)));
    const after = serialiseUnits(setUnitField(unitsOf(bulk(2)), 1, 'serialNumber', 'SN-B'));

    expect(diffUnits(before, after)).toEqual([{
      fieldName: 'Unit 2 · Serial number',
      oldValue: '',
      newValue: 'SN-B',
      changeType: 'Added',
    }]);
  });

  it('calls a cleared value removed', () => {
    const before = serialiseUnits([{ index: 0, assetTag: 'PMW-1' }]);
    const [change] = diffUnits(before, '');

    expect(change).toMatchObject({ changeType: 'Removed', oldValue: 'PMW-1', newValue: '' });
  });

  it('says nothing when nothing moved', () => {
    const stored = serialiseUnits([{ index: 0, serialNumber: 'SN-A' }]);
    expect(diffUnits(stored, stored)).toEqual([]);
  });
});

describe('counting', () => {
  it('knows how many units have been recorded', () => {
    const units = setUnitField(unitsOf(bulk(4)), 0, 'serialNumber', 'SN-A');
    expect(filledCount(units)).toBe(1);
    expect(isBlankUnit(units[1])).toBe(true);
  });
});

describe('withUnitsSplitOut', () => {
  /**
   * The rule this file exists for: a serial, a part number, a label, a
   * condition and a status each name ONE thing, so a line counted by quantity
   * does not hold them.
   */
  it('takes the per-item fields off a bulk row', () => {
    const row = {
      trackingMode: 'Bulk',
      quantity: 2,
      model: 'SM-X210',
      serialNumber: 'R52TC0AAAAA',
      partNumber: '199276824226',
      assetTag: 'PMWTAB001',
      condition: 'New',
      status: 'In stock',
    };
    const after = withUnitsSplitOut(row);

    expect(after).toMatchObject({
      model: 'SM-X210',
      serialNumber: '',
      partNumber: '',
      assetTag: '',
      condition: '',
      status: '',
    });
  });

  /** Taking them off without keeping them would be losing them. */
  it('moves what it takes onto item 1', () => {
    const after = withUnitsSplitOut({
      trackingMode: 'Bulk', quantity: 2, serialNumber: 'R52TC0AAAAA', condition: 'New',
    });

    expect(parseUnits(after.units)[0]).toMatchObject({
      index: 0, serialNumber: 'R52TC0AAAAA', condition: 'New',
    });
  });

  /**
   * A row that already has unit records has been saved under this rule. Its
   * leftovers are not a first item waiting to be adopted, and adopting them
   * would overwrite an item somebody actually recorded.
   */
  it('does not adopt row leftovers onto a row that already has items', () => {
    const stored = serialiseUnits([{ index: 0, serialNumber: 'REAL' }]);
    const after = withUnitsSplitOut({
      trackingMode: 'Bulk', quantity: 2, units: stored, serialNumber: 'LEFTOVER',
    });

    expect(after.units).toBe(stored);
    expect(after.serialNumber).toBe('');
  });

  /** A tracked row IS one item, so the row is the right place for all of it. */
  it('leaves a tracked row exactly as it found it', () => {
    const row = { trackingMode: 'Tracked', serialNumber: 'CN0ABC123', condition: 'New' };
    expect(withUnitsSplitOut(row)).toBe(row);
  });

  /**
   * A scan can claim the codes it read for the box in front of it. It cannot
   * claim a condition: "all new" on a review grid is about the delivery, and
   * writing it onto item 1 alone turns twenty new cables into one.
   */
  it('moves only the codes when that is all a scan can honestly claim', () => {
    const after = withUnitsSplitOut({
      trackingMode: 'Bulk', quantity: 20, serialNumber: 'SN-A', condition: 'New',
    }, PER_UNIT_CODES);

    expect(parseUnits(after.units)[0]).toMatchObject({ serialNumber: 'SN-A', condition: '' });
    expect(after.condition).toBe('');
  });
});

describe('mergeUnits and appendUnit', () => {
  /** A third tab arriving is a third object, not a correction to the first. */
  it('starts the arriving items after everything already on the row', () => {
    const stored = serialiseUnits([{ index: 0, serialNumber: 'SN-A' }]);
    const arriving = serialiseUnits([{ index: 0, serialNumber: 'SN-C' }]);

    expect(parseUnits(mergeUnits(stored, arriving, 2))).toEqual([
      expect.objectContaining({ index: 0, serialNumber: 'SN-A' }),
      expect.objectContaining({ index: 2, serialNumber: 'SN-C' }),
    ]);
  });

  it('steps over a position already taken rather than writing on it', () => {
    const stored = serialiseUnits([{ index: 2, serialNumber: 'SN-C' }]);
    const arriving = serialiseUnits([{ index: 0, serialNumber: 'SN-D' }]);

    expect(parseUnits(mergeUnits(stored, arriving, 2))[1]).toMatchObject({
      index: 3, serialNumber: 'SN-D',
    });
  });

  it('adds nothing for a box whose codes nobody read', () => {
    const stored = serialiseUnits([{ index: 0, serialNumber: 'SN-A' }]);
    expect(mergeUnits(stored, '', 1)).toBe(stored);
    expect(appendUnit(stored, { serialNumber: '' }, 1)).toBe(stored);
  });

  it('takes one item straight off a record that carries the fields itself', () => {
    const after = appendUnit('', { serialNumber: 'SN-A', condition: 'Good' }, 0);
    expect(parseUnits(after)[0]).toMatchObject({ serialNumber: 'SN-A', condition: 'Good' });
  });
});

describe('perItem', () => {
  /**
   * The arithmetic that stops one faulty mouse reading as twenty. Items
   * nobody has spoken for are in stock, because that is what a thing nobody
   * has handed out is.
   */
  it('counts a status across everything the row owns', () => {
    const units = serialiseUnits([{ index: 0, status: 'Retired' }]);

    expect(perItem(bulk(20, units), 'status', 'In stock')).toEqual(
      expect.arrayContaining([
        { value: 'Retired', count: 1 },
        { value: 'In stock', count: 19 },
      ]),
    );
  });

  /** A condition nobody recorded is not a condition, so it counts as nothing. */
  it('leaves unrecorded conditions out rather than inventing one', () => {
    const units = serialiseUnits([{ index: 0, condition: 'Faulty' }]);

    expect(perItem(bulk(20, units), 'condition')).toEqual([{ value: 'Faulty', count: 1 }]);
    expect(countPerItem(bulk(20, units), 'condition', 'Faulty')).toBe(1);
  });

  it('ignores items above the quantity the row claims', () => {
    const units = serialiseUnits([{ index: 5, condition: 'Faulty' }]);
    expect(countPerItem(bulk(2, units), 'condition', 'Faulty')).toBe(0);
  });

  /** A tracked row is one item and answers with its own value. */
  it('reads a tracked row off the row', () => {
    const row = { trackingMode: 'Tracked', quantity: 1, condition: 'Faulty' };
    expect(countPerItem(row, 'condition', 'Faulty')).toBe(1);
  });
});

/**
 * The pager writes every keystroke out through `serialiseUnits` and reads it
 * straight back through `unitsOf`. A trim anywhere on that path is a trim
 * applied WHILE SOMEBODY IS TYPING — and since a space is only ever typed at
 * the end of what is there so far, it made the space bar do nothing at all.
 */
describe('typing a space', () => {
  const typed = (value) => unitsOf(bulk(1), serialiseUnits([{ index: 0, location: value }]))[0];

  it('survives the round trip the pager makes on every keystroke', () => {
    expect(typed('Store ').location).toBe('Store ');
    expect(typed('Store room').location).toBe('Store room');
  });

  it('survives it however many times it is made', () => {
    let units = unitsOf(bulk(1));
    for (const value of ['S', 'St', 'Store', 'Store ', 'Store r', 'Store room']) {
      units = unitsOf(bulk(1), serialiseUnits(setUnitField(units, 0, 'location', value)));
    }
    expect(units[0].location).toBe('Store room');
  });

  /** A field holding only spaces is still nothing, and must not be stored. */
  it('does not turn a field of spaces into a filled-in item', () => {
    expect(isBlankUnit(typed('   '))).toBe(true);
    expect(serialiseUnits([{ index: 0, serialNumber: '  ' }])).toBe('');
  });
});

describe('trimUnits', () => {
  it('takes the stray spaces off on the way to storage', () => {
    const stored = serialiseUnits([{ index: 0, serialNumber: ' HA2KJDSW ' }]);
    expect(parseUnits(trimUnits(stored))[0].serialNumber).toBe('HA2KJDSW');
  });

  it('is what a save runs, so nothing half-typed reaches SharePoint', () => {
    const record = withUnitsSplitOut({
      trackingMode: 'Bulk',
      quantity: 1,
      units: serialiseUnits([{ index: 0, serialNumber: 'SN-A ', location: ' Store ' }]),
    });

    const [unit] = parseUnits(record.units);
    expect(unit.serialNumber).toBe('SN-A');
    expect(unit.location).toBe('Store');
  });
});

describe('a photograph of one item', () => {
  it('is kept on the unit, through the round trip', () => {
    const stored = serialiseUnits([{ index: 1, photoUrl: '/sites/it/photos/tab-2.jpg' }]);
    expect(parseUnits(stored)[0].photoUrl).toBe('/sites/it/photos/tab-2.jpg');
  });

  it('makes the item count as recorded, so the dots show it', () => {
    const units = unitsOf(bulk(2, serialiseUnits([{ index: 0, photoId: 'local-1' }])));
    expect(filledCount(units)).toBe(1);
  });

  /** The change log ignores photos, the same as it does for the row's own. */
  it('produces no change-log line of its own', () => {
    const before = serialiseUnits([{ index: 0, serialNumber: 'SN-A' }]);
    const after = serialiseUnits([{ index: 0, serialNumber: 'SN-A', photoUrl: '/p/a.jpg' }]);
    expect(diffUnits(before, after)).toEqual([]);
  });
});
