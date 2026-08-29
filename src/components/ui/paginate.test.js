import { describe, it, expect } from 'vitest';
import { paginate, pageCount, PAGE_SIZES } from './paginate.js';

const rows = Array.from({ length: 312 }, (unused, at) => at + 1);

describe('paginate', () => {
  it('gives one page of a long list, counted the way a reader counts', () => {
    const page = paginate(rows, 2, 25);

    expect(page.rows).toHaveLength(25);
    expect(page.rows[0]).toBe(26);
    expect(page.from).toBe(26);
    expect(page.to).toBe(50);
    expect(page.pages).toBe(13);
    expect(page.total).toBe(312);
  });

  it('answers with the last page when the list has shrunk under the reader', () => {
    const page = paginate(rows.slice(0, 30), 9, 25);

    expect(page.page).toBe(2);
    expect(page.rows).toEqual([26, 27, 28, 29, 30]);
  });

  it('hands back everything when the size is "all"', () => {
    const page = paginate(rows, 3, 0);

    expect(page.rows).toHaveLength(312);
    expect(page.pages).toBe(1);
    expect(page.page).toBe(1);
  });

  it('says nothing rather than "1 of 0" for an empty list', () => {
    const page = paginate([], 1, 25);

    expect(page.from).toBe(0);
    expect(page.to).toBe(0);
    expect(page.pages).toBe(1);
  });

  it('counts pages the way the picker offers sizes', () => {
    expect(PAGE_SIZES).toContain(25);
    expect(pageCount(312, 100)).toBe(4);
    expect(pageCount(0, 25)).toBe(1);
  });
});
