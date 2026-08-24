import { describe, it, expect } from 'vitest';
import { assetStats, recentDeliveries } from './assetStats.js';

const tracked = (overrides = {}) => ({
  category: 'Laptop', trackingMode: 'Tracked', quantity: 1, status: 'In stock', ...overrides,
});

const bulk = (overrides = {}) => ({
  category: 'Mouse', trackingMode: 'Bulk', quantity: 20, status: 'In stock', ...overrides,
});

describe('assetStats', () => {
  /** A row reading "× 20" is twenty mice, not one item. */
  it('counts units, not rows', () => {
    const stats = assetStats([tracked(), bulk()]);

    expect(stats.rows).toBe(2);
    expect(stats.units).toBe(21);
    expect(stats.trackedUnits).toBe(1);
    expect(stats.bulkUnits).toBe(20);
  });

  it('treats a row with no quantity as one unit', () => {
    expect(assetStats([{ category: 'Cable' }]).units).toBe(1);
  });

  it('counts only tracked things as needing a label', () => {
    const stats = assetStats([tracked(), tracked({ assetTag: 'PMW-1' }), bulk()]);
    expect(stats.unlabelled).toBe(1);
  });

  it('ranks the categories biggest first', () => {
    const stats = assetStats([tracked(), bulk()]);
    expect(stats.byCategory[0]).toEqual({ label: 'Mouse', value: 20 });
  });

  /** Two categories of equal size must not shuffle between renders. */
  it('breaks a tie alphabetically rather than by chance', () => {
    const stats = assetStats([
      { category: 'Zebra', quantity: 2 },
      { category: 'Apple', quantity: 2 },
    ]);
    expect(stats.byCategory.map((entry) => entry.label)).toEqual(['Apple', 'Zebra']);
  });

  it('treats a row with no status as in stock', () => {
    expect(assetStats([{ category: 'Cable', quantity: 3 }]).inStock).toBe(3);
  });

  /**
   * Condition belongs to a thing, and a bulk row is a count of things. Two of
   * twenty mice being faulty is two faulty mice — the old reading, off the
   * row, made it twenty.
   */
  it('counts the faulty ITEMS on a bulk line, not the whole line', () => {
    const units = JSON.stringify([
      { index: 0, condition: 'Faulty' },
      { index: 1, condition: 'Faulty' },
      { index: 2, condition: 'Good' },
    ]);

    expect(assetStats([bulk({ units })]).faulty).toBe(2);
  });

  it('counts a faulty tracked item once', () => {
    expect(assetStats([tracked({ condition: 'Faulty' })]).faulty).toBe(1);
  });

  it('counts nothing faulty on a line where nobody said', () => {
    expect(assetStats([bulk()]).faulty).toBe(0);
  });

  /** Items nobody has spoken for are in stock; the ones retired are not. */
  it('tallies status per item on a bulk line', () => {
    const units = JSON.stringify([{ index: 0, status: 'Retired' }]);
    const stats = assetStats([bulk({ units })]);

    expect(stats.inStock).toBe(19);
    expect(stats.byStatus).toEqual(expect.arrayContaining([{ label: 'Retired', value: 1 }]));
  });

  /**
   * A bag of cables was never going to wear twenty stickers, so a line nobody
   * has started labelling is not twenty items waiting for one. A line where
   * labelling HAS started is exactly the case worth a reminder.
   */
  it('counts the rest of a part-labelled line as waiting for a sticker', () => {
    const units = JSON.stringify([{ index: 0, assetTag: 'PMWTAB001' }]);
    const stats = assetStats([{ ...bulk({ units }), quantity: 2 }]);

    expect(stats.unlabelled).toBe(1);
  });

  it('is calm about an empty register', () => {
    expect(assetStats([])).toMatchObject({ rows: 0, units: 0, byCategory: [] });
  });
});

describe('recentDeliveries', () => {
  /** Thirty rows from one PO is one event, not thirty. */
  it('groups the rows of a delivery into one entry', () => {
    const rows = [
      tracked({ batchId: 'b1', batchTitle: 'PO-1', arrivedOn: 100 }),
      bulk({ batchId: 'b1', batchTitle: 'PO-1', arrivedOn: 100 }),
    ];
    const [delivery] = recentDeliveries(rows);

    expect(delivery.rows).toBe(2);
    expect(delivery.units).toBe(21);
  });

  it('puts the newest delivery first', () => {
    const rows = [
      tracked({ batchId: 'old', batchTitle: 'PO-1', arrivedOn: 100 }),
      tracked({ batchId: 'new', batchTitle: 'PO-2', arrivedOn: 500 }),
    ];

    expect(recentDeliveries(rows)[0].batchId).toBe('new');
  });

  it('leaves out rows that came from no delivery', () => {
    expect(recentDeliveries([tracked()])).toEqual([]);
  });

  it('stops at the limit asked for', () => {
    const rows = Array.from({ length: 9 }, (unused, index) => tracked({
      batchId: `b${index}`, batchTitle: `PO-${index}`, arrivedOn: index,
    }));

    expect(recentDeliveries(rows, 3)).toHaveLength(3);
  });
});

describe('what is out and what is left', () => {
  /** Owned stays put when stock goes out; available is the derived figure. */
  it('counts units out and units available separately from units owned', () => {
    const stats = assetStats([bulk({ quantity: 20, quantityOut: 3 })]);

    expect(stats.units).toBe(20);
    expect(stats.out).toBe(3);
    expect(stats.available).toBe(17);
  });

  it('counts a tracked item that is out as one out and none available', () => {
    const stats = assetStats([tracked({ quantityOut: 1 })]);

    expect(stats.out).toBe(1);
    expect(stats.available).toBe(0);
  });

  /** Every row saved before handovers existed has none out. */
  it('reads a row with no out-count as nothing out', () => {
    const stats = assetStats([{ category: 'Cable', trackingMode: 'Bulk', quantity: 5 }]);

    expect(stats.out).toBe(0);
    expect(stats.available).toBe(5);
  });

  it('reports nothing out for an empty register', () => {
    expect(assetStats([])).toMatchObject({ out: 0, available: 0 });
  });
});
