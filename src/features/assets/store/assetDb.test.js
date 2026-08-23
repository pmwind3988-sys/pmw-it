import { describe, it, expect, beforeEach } from 'vitest';
import 'fake-indexeddb/auto';
import {
  saveBatch, loadBatch, listBatches, deleteBatch,
  savePhoto, loadPhoto, deletePhoto, photoIdsOf, batchesBySize,
} from './assetDb.js';
import { newBatch, addDraft } from '../draft/batch.js';
import { newDraft } from '../draft/draftAsset.js';

const blob = (size) => new Blob([new Uint8Array(size)], { type: 'image/jpeg' });

beforeEach(async () => {
  for (const batch of await listBatches()) await deleteBatch(batch.id);
});

describe('batches', () => {
  it('comes back the way it went in', async () => {
    const batch = newBatch({ purchase: { supplier: 'Ingram Micro', poNumber: 'PO-1' } });
    await saveBatch(batch);

    const loaded = await loadBatch(batch.id);
    expect(loaded.purchase.supplier).toBe('Ingram Micro');
    expect(loaded.updatedAt).toEqual(expect.any(Number));
  });

  /** The whole point: a delivery scanned offline survives the app closing. */
  it('keeps the drafts scanned into it', async () => {
    const batch = addDraft(newBatch(), newDraft({ category: 'Laptop', serialNumber: 'CN0ABC' }));
    await saveBatch(batch);

    const loaded = await loadBatch(batch.id);
    expect(loaded.drafts).toHaveLength(1);
    expect(loaded.drafts[0].serialNumber).toBe('CN0ABC');
  });

  it('lists the newest delivery first', async () => {
    await saveBatch(newBatch({ createdAt: 1000 }));
    await saveBatch(newBatch({ createdAt: 3000 }));
    await saveBatch(newBatch({ createdAt: 2000 }));

    expect((await listBatches()).map((b) => b.createdAt)).toEqual([3000, 2000, 1000]);
  });

  it('returns nothing for a batch that was never saved', async () => {
    expect(await loadBatch('no-such-id')).toBeUndefined();
  });
});

describe('photos', () => {
  it('stores and returns a blob', async () => {
    await savePhoto('p1', blob(64));
    expect((await loadPhoto('p1')).size).toBe(64);
  });

  it('reads a missing photo as nothing rather than throwing', async () => {
    expect(await loadPhoto('p-nope')).toBeNull();
    expect(await loadPhoto(null)).toBeNull();
  });

  it('deletes one', async () => {
    await savePhoto('p1', blob(8));
    await deletePhoto('p1');
    expect(await loadPhoto('p1')).toBeNull();
  });
});

describe('deleting a batch', () => {
  /**
   * A photo left behind occupies storage nothing can reach or free — the one
   * leak a quota-limited store never recovers from.
   */
  it('takes every photo belonging to it', async () => {
    const draft = newDraft({ photoId: 'item-photo' });
    const batch = addDraft(
      newBatch({ purchase: { poPhotoId: 'po-photo' } }),
      draft,
    );
    await savePhoto('item-photo', blob(16));
    await savePhoto('po-photo', blob(16));
    await saveBatch(batch);

    await deleteBatch(batch.id);

    expect(await loadPhoto('item-photo')).toBeNull();
    expect(await loadPhoto('po-photo')).toBeNull();
    expect(await loadBatch(batch.id)).toBeUndefined();
  });

  it('names both the item photos and the PO scan', () => {
    const batch = addDraft(
      newBatch({ purchase: { poPhotoId: 'po' } }),
      newDraft({ photoId: 'a' }),
    );

    expect(photoIdsOf(batch).sort()).toEqual(['a', 'po']);
  });

  it('is safe on a batch that does not exist', async () => {
    await expect(deleteBatch('no-such-id')).resolves.toBeUndefined();
  });
});

describe('batchesBySize', () => {
  it('puts the delivery taking the most room first', async () => {
    const small = addDraft(newBatch({ createdAt: 2 }), newDraft({ photoId: 's1' }));
    const large = addDraft(newBatch({ createdAt: 1 }), newDraft({ photoId: 'l1' }));
    await savePhoto('s1', blob(100));
    await savePhoto('l1', blob(5000));
    await saveBatch(small);
    await saveBatch(large);

    const sized = await batchesBySize();
    expect(sized[0].bytes).toBe(5000);
    expect(sized[1].bytes).toBe(100);
  });

  it('reports zero for a delivery with no photos at all', async () => {
    await saveBatch(newBatch());
    expect((await batchesBySize())[0].bytes).toBe(0);
  });
});
