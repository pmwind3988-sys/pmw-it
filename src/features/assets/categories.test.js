import { describe, it, expect } from 'vitest';
import { categoriesIn, categoryRefusal, cleanCategory } from './categories.js';
import { CATEGORIES } from './assetKinds.js';

describe('categoriesIn', () => {
  it('offers the built-in list when nothing else is in use', () => {
    expect(categoriesIn([])).toEqual(CATEGORIES);
  });

  it('offers a category somebody added, once, after the built-in ones', () => {
    const list = categoriesIn([
      { category: 'Projector' }, { category: 'Projector' }, { category: 'Laptop' },
    ]);

    expect(list).toContain('Projector');
    expect(list.filter((name) => name === 'Projector')).toHaveLength(1);
    expect(list.indexOf('Projector')).toBe(CATEGORIES.length);
  });

  it('does not offer the same name twice in two spellings', () => {
    expect(categoriesIn([{ category: 'laptop' }])).toEqual(CATEGORIES);
  });
});

describe('categoryRefusal', () => {
  it('accepts something new', () => {
    expect(categoryRefusal('Projector')).toBe('');
  });

  it('refuses a blank, a duplicate and a remark', () => {
    expect(categoryRefusal('   ')).toMatch(/Type a name/);
    expect(categoryRefusal('  laptop ')).toMatch(/already on the list/);
    expect(categoryRefusal('x'.repeat(61))).toMatch(/too long/);
  });

  it('tidies the spacing rather than refusing it', () => {
    expect(cleanCategory('  Docking   Station ')).toBe('Docking Station');
  });
});
