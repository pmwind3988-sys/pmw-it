import { describe, it, expect } from 'vitest';
import { filterAssets, matchesQuery, sortAssets, optionsFor } from './assetFilters.js';

const laptop = {
  title: 'Dell Latitude 5540 — CN0ABC123',
  category: 'Laptop',
  trackingMode: 'Tracked',
  manufacturer: 'Dell',
  model: 'Latitude 5540',
  serialNumber: 'CN0ABC123',
  assetTag: 'PMW-0142',
  status: 'In stock',
  condition: 'New',
  location: 'Store room',
  additionalCodes: ['X99Z'],
  quantity: 1,
  arrivedOn: 500,
};

const mice = {
  title: 'Logitech B100',
  category: 'Mouse',
  trackingMode: 'Bulk',
  manufacturer: 'Logitech',
  model: 'B100',
  status: 'In stock',
  condition: 'New',
  quantity: 20,
  arrivedOn: 100,
};

describe('matchesQuery', () => {
  it('finds a row by its model', () => {
    expect(matchesQuery(laptop, 'latitude')).toBe(true);
  });

  it('finds a row by its serial number', () => {
    expect(matchesQuery(laptop, 'CN0ABC123')).toBe(true);
  });

  it('finds a row by its sticker label', () => {
    expect(matchesQuery(laptop, 'pmw-0142')).toBe(true);
  });

  /** A code nobody could place is still how somebody might look for the thing. */
  it('reaches the codes that went into Other Codes', () => {
    expect(matchesQuery(laptop, 'x99z')).toBe(true);
  });

  /** The label reads CN0ABC123; a person types it with a space in it. */
  it('ignores spacing in something that looks like a code', () => {
    expect(matchesQuery(laptop, 'cn0abc 123')).toBe(true);
  });

  it('matches everything when nothing was typed', () => {
    expect(matchesQuery(mice, '')).toBe(true);
    expect(matchesQuery(mice, '   ')).toBe(true);
  });

  it('says no when it genuinely is not there', () => {
    expect(matchesQuery(mice, 'thinkpad')).toBe(false);
  });
});

describe('filterAssets', () => {
  const all = [laptop, mice];

  it('shows everything when no filter is set', () => {
    expect(filterAssets(all, {})).toHaveLength(2);
    expect(filterAssets(all, { category: '', status: '' })).toHaveLength(2);
  });

  it('narrows by category', () => {
    expect(filterAssets(all, { category: 'Mouse' })).toEqual([mice]);
  });

  it('narrows by tracking mode', () => {
    expect(filterAssets(all, { trackingMode: 'Tracked' })).toEqual([laptop]);
  });

  it('finds what still needs a sticker', () => {
    expect(filterAssets(all, { unlabelled: true })).toEqual([mice]);
  });

  it('treats a missing status as in stock rather than hiding the row', () => {
    expect(filterAssets([{ category: 'Cable' }], { status: 'In stock' })).toHaveLength(1);
  });

  it('applies the search and the filters together', () => {
    expect(filterAssets(all, { category: 'Laptop', query: 'logitech' })).toEqual([]);
  });
});

describe('sortAssets', () => {
  it('puts the newest arrival first by default', () => {
    expect(sortAssets([mice, laptop])[0]).toBe(laptop);
  });

  /** A row saved before the column existed is not the newest thing here. */
  it('sinks rows with no arrival date', () => {
    const undated = { title: 'Old', quantity: 1 };
    expect(sortAssets([undated, mice]).at(-1)).toBe(undated);
  });

  it('sorts by name and by quantity when asked', () => {
    expect(sortAssets([mice, laptop], 'name')[0]).toBe(laptop);
    expect(sortAssets([laptop, mice], 'quantity')[0]).toBe(mice);
  });

  it('does not disturb the list it was given', () => {
    const input = [mice, laptop];
    sortAssets(input, 'name');
    expect(input[0]).toBe(mice);
  });
});

describe('optionsFor', () => {
  it('offers only the values actually present, in order', () => {
    expect(optionsFor([laptop, mice], 'category')).toEqual(['Laptop', 'Mouse']);
  });

  it('leaves out blanks', () => {
    expect(optionsFor([laptop, mice], 'location')).toEqual(['Store room']);
  });
});
