import { describe, it, expect } from 'vitest';
import { newBatch } from '../draft/batch.js';
import { TRACKED, BULK } from '../assetKinds.js';
import {
  scanIn, addModel, STOCK_RESULT, matchRegister, needsKind, needsSerial, setKind, nextTag,
  linkToModel, addSerialsTo, serialCount,
  noCodeDraft, holdsFor, itemCount,
} from './stockIn.js';

const LAPTOP = {
  id: 1, assetKey: 'serial:DELL|5CG4211XQ7', category: 'Laptop', trackingMode: TRACKED,
  manufacturer: 'Dell', model: 'Latitude 5450', serialNumber: '5CG4211XQ7',
  partNumber: '0X7K3M', assetTag: 'PMW-0142', location: 'F1', quantity: 1,
};
const MICE = {
  id: 2, assetKey: 'bulk:Mouse|LOGITECH|M90', category: 'Mouse', trackingMode: BULK,
  manufacturer: 'Logitech', model: 'M90', partNumber: '5099206092372', location: 'IT store', quantity: 17,
};
const REGISTER = [LAPTOP, MICE];

describe('matchRegister', () => {
  it('tells a serial or label from a part number', () => {
    expect(matchRegister(REGISTER, '5cg4211xq7').by).toBe('serial');
    expect(matchRegister(REGISTER, 'pmw-0142').by).toBe('tag');
    expect(matchRegister(REGISTER, '0X7K3M').by).toBe('part');
    expect(matchRegister(REGISTER, 'nothing')).toBeNull();
  });
});

describe('scanIn', () => {
  it('names a bulk model the register knows and counts the next bag up', () => {
    let { batch, result, draft } = scanIn(newBatch(), '5099206092372', REGISTER);
    expect(result).toBe(STOCK_RESULT.ADDED);
    expect(draft).toMatchObject({ category: 'Mouse', model: 'M90', quantity: 1, location: 'IT store' });

    ({ batch, result, draft } = scanIn(batch, '5099206092372', REGISTER));
    expect(result).toBe(STOCK_RESULT.COUNTED);
    expect(batch.drafts).toHaveLength(1);
    expect(draft.quantity).toBe(2);
  });

  it('refuses to count a tracked box read twice', () => {
    let { batch } = scanIn(newBatch(), 'ABC123XYZ9', []);
    const again = scanIn(batch, 'abc123xyz9', []);
    expect(again.result).toBe(STOCK_RESULT.DUPLICATE);
    expect(again.batch.drafts).toHaveLength(1);
  });

  it('leaves a machine that is already registered alone', () => {
    const { batch, result, asset } = scanIn(newBatch(), '5CG4211XQ7', REGISTER);
    expect(result).toBe(STOCK_RESULT.KNOWN);
    expect(asset).toBe(LAPTOP);
    expect(batch.drafts).toHaveLength(0);
  });

  it('reads a tracked part number as another one of that model, waiting for its serial', () => {
    let { batch, result, draft } = scanIn(newBatch(), '0X7K3M', REGISTER);
    expect(result).toBe(STOCK_RESULT.ADDED);
    expect(draft).toMatchObject({ model: 'Latitude 5450', trackingMode: TRACKED, serialNumber: '' });
    expect(needsSerial(draft)).toBe(true);

    ({ batch, result, draft } = scanIn(batch, '5CG4211ZZ1', REGISTER));
    expect(result).toBe(STOCK_RESULT.FILLED);
    expect(batch.drafts).toHaveLength(1);
    expect(draft.serialNumber).toBe('5CG4211ZZ1');
  });

  it('adds an unknown code as a line that asks what it is', () => {
    const { batch, result, draft } = scanIn(newBatch(), '4710886112233', REGISTER);
    expect(result).toBe(STOCK_RESULT.NEW);
    expect(needsKind(draft)).toBe(true);
    expect(holdsFor(batch, REGISTER).get(draft.localId)).toBe('Say what this is.');

    const named = setKind(batch, draft.localId, 'Cable');
    expect(needsKind(named.drafts[0])).toBe(false);
  });

  it('ignores an empty read', () => {
    expect(scanIn(newBatch(), '  ', REGISTER).result).toBe(STOCK_RESULT.EMPTY);
  });
});

describe('addModel', () => {
  it('counts a picked bulk model up and adds a picked tracked one waiting for its serial', () => {
    let { batch } = addModel(newBatch(), MICE);
    let result;
    ({ batch, result } = addModel(batch, MICE));
    expect(result).toBe(STOCK_RESULT.COUNTED);
    expect(batch.drafts[0].quantity).toBe(2);

    ({ batch } = addModel(batch, LAPTOP));
    expect(batch.drafts).toHaveLength(2);
    expect(needsSerial(batch.drafts[1])).toBe(true);
    expect(batch.drafts[1].serialNumber).toBe('');
  });
});

describe('nextTag', () => {
  it('counts on from the highest label in the register and on the receipt', () => {
    expect(nextTag(REGISTER, [])).toBe('PMW-0143');
    expect(nextTag(REGISTER, [{ assetTag: 'PMW-0200' }])).toBe('PMW-0201');
    expect(nextTag([], [])).toBe('PMW-0001');
  });
});

describe('noCodeDraft', () => {
  it('gives a tracked thing its label and a bulk thing its count', () => {
    expect(noCodeDraft({ category: 'Laptop', model: 'X1', tag: 'PMW-0143', quantity: 5 }))
      .toMatchObject({ assetTag: 'PMW-0143', quantity: 1, trackingMode: TRACKED });
    expect(noCodeDraft({ category: 'Cable', model: 'HDMI 1.8 m', tag: 'PMW-0143', quantity: 5 }))
      .toMatchObject({ assetTag: '', quantity: 5, trackingMode: BULK });
  });
});

describe('itemCount', () => {
  it('counts items, not lines', () => {
    let { batch } = scanIn(newBatch(), '5099206092372', REGISTER);
    ({ batch } = scanIn(batch, '5099206092372', REGISTER));
    ({ batch } = scanIn(batch, '0X7K3M', REGISTER));
    expect(itemCount(batch)).toBe(3);
  });
});

describe('new stock of a model already counted', () => {
  // A counted line as it comes back from SharePoint: its box barcode was
  // filed under item 1, and one more barcode remembered on the line.
  const SAVED_MICE = {
    ...MICE, partNumber: '', additionalCodes: ['MOUSEBOX2'],
    units: JSON.stringify([{ index: 0, partNumber: '5099206092372', serialNumber: '2133LZ0A1' }]),
  };

  it('recognises the box by the barcode on its first item, or one remembered on the line', () => {
    expect(matchRegister([SAVED_MICE], '5099206092372')).toMatchObject({ by: 'part' });
    expect(matchRegister([SAVED_MICE], 'mousebox2')).toMatchObject({ by: 'part' });
    expect(matchRegister([SAVED_MICE], '2133LZ0A1')).toMatchObject({ by: 'serial' });
  });

  it('counts the new boxes onto one line that remembers the barcode', () => {
    let { batch, draft } = scanIn(newBatch(), '5099206092372', [SAVED_MICE]);
    expect(draft.additionalCodes).toEqual(['MOUSEBOX2', '5099206092372']);
    expect(draft.partNumber).toBe('');
    ({ draft } = scanIn(batch, '5099206092372', [SAVED_MICE]));
    expect(draft.quantity).toBe(2);
  });

  it('turns an unknown code into a model we have, and counts the next box on sight', () => {
    let { batch, draft } = scanIn(newBatch(), 'NEWBOX123', [SAVED_MICE]);
    batch = linkToModel(batch, draft.localId, SAVED_MICE);
    expect(batch.drafts[0]).toMatchObject({ model: 'M90', manufacturer: 'Logitech', quantity: 1 });
    expect(batch.drafts[0].additionalCodes).toContain('NEWBOX123');
    const again = scanIn(batch, 'NEWBOX123', [SAVED_MICE]);
    expect(again.result).toBe(STOCK_RESULT.COUNTED);
    expect(again.batch.drafts).toHaveLength(1);
  });

  it('counts a box with the model’s other barcode onto the same line', () => {
    let { batch, draft } = scanIn(newBatch(), 'NEWBOX123', [SAVED_MICE]);
    batch = linkToModel(batch, draft.localId, SAVED_MICE);
    const out = scanIn(batch, '5099206092372', [SAVED_MICE]);
    expect(out.result).toBe(STOCK_RESULT.COUNTED);
    expect(out.batch.drafts).toHaveLength(1);
    expect(out.draft.quantity).toBe(2);
    expect(out.draft.additionalCodes).toEqual(expect.arrayContaining(['NEWBOX123', '5099206092372']));
  });

  it('links an unknown laptop code as its serial', () => {
    const { batch, draft } = scanIn(newBatch(), '5CG9999ZZ1', []);
    const linked = linkToModel(batch, draft.localId, LAPTOP);
    expect(linked.drafts[0]).toMatchObject({ model: 'Latitude 5450', serialNumber: '5CG9999ZZ1' });
  });

  it('puts a serial run onto a counted line, one item each', () => {
    let { batch, draft } = scanIn(newBatch(), '5099206092372', REGISTER);
    batch = addSerialsTo(batch, draft.localId, ['S1', 'S2', 'S3'], 1);
    expect(batch.drafts[0].quantity).toBe(4);
    expect(serialCount(batch.drafts[0])).toBe(3);
  });
});
