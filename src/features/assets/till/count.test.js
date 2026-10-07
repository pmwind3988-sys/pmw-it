import { describe, it, expect } from 'vitest';
import { TRACKED, BULK } from '../assetKinds.js';
import { parseUnits } from '../units.js';
import { holdsFor, itemCount } from './stockIn.js';
import {
  newCount, addCounted, recentModels, knownLocations, COUNT_RESULT, COUNT_REMARK,
} from './count.js';

const REGISTER = [
  { id: 1, category: 'Monitor', trackingMode: TRACKED, manufacturer: 'Dell', model: 'P2422H', serialNumber: 'CN0WJ2K1', location: 'F1' },
  { id: 2, category: 'Monitor', trackingMode: TRACKED, manufacturer: 'Dell', model: 'P2422H', serialNumber: 'CN0WJ2K2', location: 'F1' },
  { id: 3, category: 'Monitor', trackingMode: TRACKED, manufacturer: 'LG', model: '24MK430H', serialNumber: 'LG1', location: 'F3' },
  { id: 4, category: 'Mouse', trackingMode: BULK, manufacturer: 'Logitech', model: 'M90', quantity: 17, location: 'IT store' },
];

describe('newCount', () => {
  it('is a batch with no arrival date to invent', () => {
    const count = newCount();
    expect(count.purchase.arrivedOn).toBeNull();
    expect(count.drafts).toEqual([]);
  });
});

describe('addCounted', () => {
  it('counts a drawer of mice as one line, and a second drawer in the same place onto it', () => {
    let { batch, result, draft } = addCounted(newCount(), { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', quantity: 14, location: 'F1' });
    expect(result).toBe(COUNT_RESULT.ADDED);
    expect(draft).toMatchObject({ trackingMode: BULK, quantity: 14, location: 'F1', remarks: COUNT_REMARK });

    ({ batch, result } = addCounted(batch, { category: 'Mouse', manufacturer: 'logitech', model: 'm90', quantity: 3, location: 'F1' }));
    expect(result).toBe(COUNT_RESULT.COUNTED);
    expect(batch.drafts).toHaveLength(1);
    expect(batch.drafts[0].quantity).toBe(17);

    ({ batch } = addCounted(batch, { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', quantity: 2, location: 'F3' }));
    expect(batch.drafts).toHaveLength(2);
    expect(itemCount(batch)).toBe(19);
  });

  it('adds one line per tracked thing, with its serial', () => {
    const { batch, draft } = addCounted(newCount(), { category: 'Monitor', manufacturer: 'Dell', model: 'P2422H', serialNumber: 'cn0zz9', condition: 'Fair' });
    expect(draft).toMatchObject({ serialNumber: 'CN0ZZ9', condition: 'Fair', trackingMode: TRACKED });
    expect(holdsFor(batch).size).toBe(0);
    expect(addCounted(batch, { category: 'Monitor', model: 'P2422H', serialNumber: 'CN0ZZ9' }).result).toBe(COUNT_RESULT.DUPLICATE);
  });

  it('refuses a serial the register already has', () => {
    const out = addCounted(newCount(), { category: 'Monitor', model: 'P2422H', serialNumber: 'CN0WJ2K1' }, REGISTER);
    expect(out.result).toBe(COUNT_RESULT.KNOWN);
    expect(out.asset.id).toBe(1);
  });

  it('takes a tracked thing with no serial only when told so, and does not hold the count for it', () => {
    expect(addCounted(newCount(), { category: 'Printer', model: 'HP M404' }).result).toBe(COUNT_RESULT.INVALID);

    const { batch, draft } = addCounted(newCount(), { category: 'Printer', model: 'HP M404', noSerial: true, photoId: 'p1' });
    expect(draft).toMatchObject({ noSerial: true, photoId: 'p1', serialNumber: '' });
    expect(draft.remarks).toMatch(/No serial/);
    expect(holdsFor(batch).size).toBe(0);
  });

  it('needs a category and a model', () => {
    expect(addCounted(newCount(), { category: 'Mouse', model: ' ' }).result).toBe(COUNT_RESULT.INVALID);
  });
});

describe('recentModels', () => {
  it('offers what this count just used first, then what the register holds most of', () => {
    const drafts = [{ category: 'Monitor', manufacturer: 'LG', model: '24MK430H' }];
    expect(recentModels('Monitor', REGISTER, drafts)).toEqual([
      { manufacturer: 'LG', model: '24MK430H' },
      { manufacturer: 'Dell', model: 'P2422H' },
    ]);
    expect(recentModels('Mouse', REGISTER)).toEqual([{ manufacturer: 'Logitech', model: 'M90' }]);
  });
});

describe('knownLocations', () => {
  it('lists each place once, the count’s own first', () => {
    expect(knownLocations(REGISTER, [{ location: 'F3' }])).toEqual(['F3', 'F1', 'IT store']);
  });
});

describe('a serial run in a count', () => {
  it('makes one item per serial, and a later run lands after the first', () => {
    let { batch, draft } = addCounted(newCount(), { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', location: 'F1', serials: ['A1', 'A2'], without: 1 });
    expect(draft.quantity).toBe(3);
    ({ draft } = addCounted(batch, { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', location: 'F1', serials: ['B1'] }));
    expect(draft.quantity).toBe(4);
    expect(parseUnits(draft.units).map((unit) => [unit.index, unit.serialNumber])).toEqual([[0, 'A1'], [1, 'A2'], [3, 'B1']]);
  });
});

describe('details read off the box', () => {
  it('go with the counted line, and a later count keeps the first ones', () => {
    let { batch, draft } = addCounted(newCount(), { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', location: 'F1', quantity: 2, specSummary: 'Black · Wired' });
    expect(draft.specSummary).toBe('Black · Wired');
    ({ draft } = addCounted(batch, { category: 'Mouse', manufacturer: 'Logitech', model: 'M90', location: 'F1', quantity: 1, specSummary: '' }));
    expect(draft.specSummary).toBe('Black · Wired');
  });
});
