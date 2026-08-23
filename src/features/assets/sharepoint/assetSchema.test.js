import { describe, it, expect } from 'vitest';
import {
  ASSET_COLUMNS, BATCH_COLUMNS, CHANGE_COLUMNS, toListItem, fromListItem, toUpdateItem,
  ASSET_LIST_NAME, BATCH_LIST_NAME, CHANGE_LIST_NAME,
} from './assetSchema.js';
import { ASSET_VIEWS } from './assetViews.js';
import { HANDOVER_COLUMNS, HANDOVER_LIST_NAME } from './handoverSchema.js';

describe('the column declaration', () => {
  it('names every column exactly once', () => {
    for (const columns of [ASSET_COLUMNS, BATCH_COLUMNS, CHANGE_COLUMNS, HANDOVER_COLUMNS]) {
      const names = columns.map((column) => column.StaticName);
      expect(new Set(names).size).toBe(names.length);
    }
  });

  /**
   * SharePoint derives the internal name from the Title a field is CREATED
   * with, so an internal name carrying a space becomes `Foo_x0020_Bar` and
   * every item write of it then fails. The provisioner creates under
   * StaticName; this pins that StaticName is safe to create under.
   */
  it('uses internal names that survive being created', () => {
    for (const columns of [ASSET_COLUMNS, BATCH_COLUMNS, CHANGE_COLUMNS, HANDOVER_COLUMNS]) {
      for (const column of columns) {
        expect(column.StaticName).toMatch(/^[A-Za-z][A-Za-z0-9]*$/);
      }
    }
  });

  it('gives every choice column its choices', () => {
    for (const column of ASSET_COLUMNS.filter((c) => c.kind === 'choice')) {
      expect(column.choices?.length).toBeGreaterThan(0);
    }
  });

  /**
   * A view can only show a column that exists — a view naming a column that
   * does not fails provisioning outright, and does so at the very end of a
   * minute-long first run.
   *
   * Reported as a list of names rather than a bare boolean: "expected false to
   * be true" does not say which column, and this test exists precisely for the
   * moment somebody adds a view faster than they add its column.
   */
  it('only shows columns the lists actually have', () => {
    const known = new Set([
      'LinkTitle', 'Title',
      ...ASSET_COLUMNS.map((c) => c.StaticName),
      ...BATCH_COLUMNS.map((c) => c.StaticName),
      ...CHANGE_COLUMNS.map((c) => c.StaticName),
      ...HANDOVER_COLUMNS.map((c) => c.StaticName),
    ]);

    const unknown = ASSET_VIEWS.flatMap(
      (view) => view.fields
        .filter((field) => !known.has(field))
        .map((field) => `${view.list} / ${view.title}: ${field}`),
    );

    expect(unknown).toEqual([]);
  });

  /** Every view must name a list that is actually provisioned. */
  it('only names lists that are provisioned', () => {
    const lists = new Set([
      ASSET_LIST_NAME, BATCH_LIST_NAME, CHANGE_LIST_NAME, HANDOVER_LIST_NAME,
    ]);

    expect(ASSET_VIEWS.filter((view) => !lists.has(view.list))).toEqual([]);
  });
});

describe('toListItem', () => {
  const asset = {
    title: 'Dell Latitude 5540 — CN0ABC123',
    assetKey: 'serial:DELL|CN0ABC123',
    category: 'Laptop',
    trackingMode: 'Tracked',
    manufacturer: 'Dell',
    serialNumber: 'CN0ABC123',
    quantity: 1,
    arrivedOn: 1755950400000,
    addedOn: 1755950400000,
    additionalCodes: ['X1', 'Y2'],
  };

  it('sends the name as Title and the identity as its own column', () => {
    const item = toListItem(asset);
    expect(item.Title).toBe('Dell Latitude 5540 — CN0ABC123');
    expect(item.AssetKey).toBe('serial:DELL|CN0ABC123');
  });

  it('sends a date as an ISO instant and its readable twin beside it', () => {
    const item = toListItem(asset);
    expect(item.ArrivedOn).toBe(new Date(1755950400000).toISOString());
    expect(item.ArrivedOnMYT).toMatch(/(AM|PM)/);
  });

  it('joins a list of codes into lines rather than sending "[object Object]"', () => {
    expect(toListItem(asset).AdditionalCodes).toBe('X1\nY2');
  });

  /** Empty string clears a column; null would be rejected outright. */
  it('clears a text column rather than omitting it', () => {
    expect(toListItem({ ...asset, location: null }).Location).toBe('');
  });

  it('omits a number that is not one, rather than sending NaN', () => {
    const item = toListItem({ ...asset, quantity: undefined });
    expect('Quantity' in item).toBe(false);
  });

  it('omits an unparseable date rather than sending "Invalid Date"', () => {
    const item = toListItem({ ...asset, purchasedOn: 'last Tuesday' });
    expect('PurchasedOn' in item).toBe(false);
  });
});

describe('fromListItem', () => {
  it('reads a row back into the shape the app works in', () => {
    const record = fromListItem({
      Id: 7,
      Title: 'Dell Latitude 5540',
      AssetKey: 'serial:DELL|CN0ABC123',
      Category: 'Laptop',
      Quantity: 3,
      ArrivedOn: '2026-08-23T04:00:00Z',
      AdditionalCodes: 'X1\nY2',
    });

    expect(record).toMatchObject({
      id: 7, title: 'Dell Latitude 5540', category: 'Laptop', quantity: 3,
    });
    expect(record.arrivedOn).toBe(Date.parse('2026-08-23T04:00:00Z'));
    expect(record.additionalCodes).toEqual(['X1', 'Y2']);
  });

  /** `new Date(undefined).getTime()` is NaN, which then poisons every sort. */
  it('reads an absent date as null, not as NaN', () => {
    expect(fromListItem({ Id: 1 }).arrivedOn).toBeNull();
  });

  it('reads an absent list column as an empty list', () => {
    expect(fromListItem({ Id: 1 }).additionalCodes).toEqual([]);
  });

  /** Every count on the page sums this; a null would silently drop a row. */
  it('defaults a missing quantity to one', () => {
    expect(fromListItem({ Id: 1 }).quantity).toBe(1);
  });

  it('round-trips an asset through both directions unchanged', () => {
    const asset = {
      title: 'Logitech B100',
      assetKey: 'bulk:MOUSE|LOGITECH|B100',
      category: 'Mouse',
      trackingMode: 'Bulk',
      manufacturer: 'Logitech',
      model: 'B100',
      quantity: 20,
      condition: 'New',
      status: 'In stock',
      location: 'Store room',
    };
    const back = fromListItem({ Id: 1, ...toListItem(asset) });

    expect(back).toMatchObject({
      title: asset.title,
      assetKey: asset.assetKey,
      category: 'Mouse',
      quantity: 20,
      location: 'Store room',
    });
  });
});

describe('toUpdateItem', () => {
  /**
   * The whole reason this exists apart from `toListItem`. A handover setting
   * `quantityOut` through `toListItem` would send an empty Serial Number,
   * Supplier and Photo for every item it touched, because a record built from
   * a patch has nothing in the rest.
   */
  it('sends only what it was given', () => {
    const item = toUpdateItem({ quantityOut: 3 });

    expect(item).toEqual({ QuantityOut: 3 });
    expect('SerialNumber' in item).toBe(false);
    expect('Supplier' in item).toBe(false);
  });

  it('sends zero, which is a real number of things out', () => {
    expect(toUpdateItem({ quantityOut: 0 })).toEqual({ QuantityOut: 0 });
  });

  /** How a return stops the row advertising a deadline that has passed. */
  it('clears a date by sending null', () => {
    expect(toUpdateItem({ dueOn: null })).toEqual({ DueOn: null });
  });

  it('writes a date as an ISO instant', () => {
    expect(toUpdateItem({ assignedOn: 1755950400000 }))
      .toEqual({ AssignedOn: new Date(1755950400000).toISOString() });
  });

  /** A choice column accepts null to clear it, but not an empty string. */
  it('clears a choice with null rather than with an empty string', () => {
    expect(toUpdateItem({ handoverKind: '' })).toEqual({ HandoverKind: null });
  });

  it('clears a text column with an empty string', () => {
    expect(toUpdateItem({ assignedTo: '' })).toEqual({ AssignedTo: '' });
  });

  it('ignores a key that is not a column', () => {
    expect(toUpdateItem({ notAColumn: 'x' })).toEqual({});
  });
});
