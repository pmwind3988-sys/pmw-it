import { describe, it, expect } from 'vitest';
import {
  parseUnits, unitsOf, setUnitField, serialiseUnits, diffUnits, filledCount, isBlankUnit,
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
