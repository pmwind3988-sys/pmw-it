import { describe, it, expect } from 'vitest';
import {
  newBatch, resolveDraft, resolveDrafts, batchTitle,
  addDraft, replaceDraft, removeDraft, mergeDrafts,
} from './batch.js';
import { newDraft, draftFromCodes } from './draftAsset.js';

const delivery = () => newBatch({
  purchase: { supplier: 'Ingram Micro', poNumber: 'PO-4471', arrivedOn: 1755950400000 },
});

describe('newBatch', () => {
  it('keeps the rest of the purchase details when only one is given', () => {
    const batch = newBatch({ purchase: { supplier: 'Ingram Micro' } });

    expect(batch.supplier).toBeUndefined();
    expect(batch.purchase.supplier).toBe('Ingram Micro');
    expect(batch.purchase.arrivedOn).toEqual(expect.any(Number));
  });
});

describe('purchase details flowing down', () => {
  it('fills a row from the delivery', () => {
    const resolved = resolveDraft(newDraft(), delivery());

    expect(resolved.supplier).toBe('Ingram Micro');
    expect(resolved.poNumber).toBe('PO-4471');
    expect(resolved.arrivedOn).toBe(1755950400000);
  });

  /**
   * The case this distinction exists for: one line on the delivery note came
   * from somewhere else, and the other rows must not follow it.
   */
  it('lets one row override the supplier without disturbing the others', () => {
    const batch = delivery();
    const withRows = addDraft(
      addDraft(batch, newDraft({ supplier: 'Lazada' })),
      newDraft(),
    );
    const [odd, normal] = resolveDrafts(withRows);

    expect(odd.supplier).toBe('Lazada');
    expect(normal.supplier).toBe('Ingram Micro');
  });

  it('treats an empty string as a deliberate "no supplier", not as inherit', () => {
    expect(resolveDraft(newDraft({ supplier: '' }), delivery()).supplier).toBe('');
  });

  it('stamps every row with which delivery it came from', () => {
    const batch = delivery();
    expect(resolveDraft(newDraft(), batch).batchId).toBe(batch.id);
    expect(resolveDraft(newDraft(), batch).batchTitle).toBe('PO-4471');
  });
});

describe('batchTitle', () => {
  it('is the PO number when there is one, without stuttering the prefix', () => {
    expect(batchTitle(delivery())).toBe('PO-4471');
  });

  it('falls back to the supplier and the day', () => {
    const batch = newBatch({ purchase: { supplier: 'Ingram Micro', arrivedOn: 1755950400000 } });
    expect(batchTitle(batch)).toBe('Ingram Micro — 2025-08-23');
  });

  it('never comes back blank', () => {
    expect(batchTitle({ purchase: {}, createdAt: Date.now() })).not.toBe('');
  });
});

describe('editing the batch', () => {
  it('adds, replaces and removes rows', () => {
    const draft = newDraft();
    const batch = addDraft(delivery(), draft);

    expect(batch.drafts).toHaveLength(1);
    expect(replaceDraft(batch, { ...draft, model: 'X' }).drafts[0].model).toBe('X');
    expect(removeDraft(batch, draft.localId).drafts).toHaveLength(0);
  });
});

describe('merging two rows that turned out to be one box', () => {
  const twoHalves = () => {
    const serialRow = draftFromCodes([{ rawValue: 'CN0ABC1234567' }]);
    const partRow = draftFromCodes([{ rawValue: '5901234123457', format: 'ean_13' }]);
    return { serialRow, partRow, batch: addDraft(addDraft(delivery(), serialRow), partRow) };
  };

  it('leaves one row carrying both codes', () => {
    const { serialRow, partRow, batch } = twoHalves();
    const merged = mergeDrafts(batch, serialRow.localId, partRow.localId);

    expect(merged.drafts).toHaveLength(1);
    expect(merged.drafts[0].serialNumber).toBe('CN0ABC1234567');
    expect(merged.drafts[0].partNumber).toBe('5901234123457');
  });

  it('keeps what the kept row already had', () => {
    const { serialRow, partRow, batch } = twoHalves();
    const withModel = replaceDraft(batch, { ...serialRow, model: 'Latitude 5540' });
    const merged = mergeDrafts(withModel, serialRow.localId, partRow.localId);

    expect(merged.drafts[0].model).toBe('Latitude 5540');
  });

  /** A code nobody can place is still the only copy of what was on the box. */
  it('never drops a code the absorbed row carried', () => {
    const { serialRow, partRow, batch } = twoHalves();
    const withSerial = replaceDraft(batch, { ...partRow, serialNumber: 'OTHER-999' });
    const merged = mergeDrafts(withSerial, serialRow.localId, partRow.localId);

    expect(merged.drafts[0].additionalCodes).toContain('OTHER-999');
  });

  it('does nothing when either row is missing, or they are the same row', () => {
    const { serialRow, batch } = twoHalves();

    expect(mergeDrafts(batch, serialRow.localId, 'nope')).toBe(batch);
    expect(mergeDrafts(batch, serialRow.localId, serialRow.localId)).toBe(batch);
  });
});
