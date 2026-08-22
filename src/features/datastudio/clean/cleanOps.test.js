import { describe, it, expect } from 'vitest';
import {
  trimWhitespace, normalizeNulls, parseNumber, unifyCase,
  categoryKey, clusterCategories, mergeCategories, dedupeRows, dropEmptyColumns,
  dropEmptyRows, parseDate, castType,
} from './cleanOps.js';

describe('trimWhitespace', () => {
  it('trims, collapses internal runs, and strips non-breaking spaces', () => {
    // The NBSP is an escape sequence on purpose: a literal U+00A0 is
    // invisible in a diff and does not survive being retyped, which
    // would quietly turn the claim in this test's name into a no-op.
    expect(trimWhitespace(['  a  b  ', 'c\u00A0d']))
      .toEqual(['a b', 'c d']);
  });

  it('leaves non-strings alone', () => {
    const d = new Date();
    expect(trimWhitespace([1, d])).toEqual([1, d]);
  });

  it('strips zero-width characters, which the whitespace class does not cover', () => {
    expect(trimWhitespace(['a\u200Bb', '\uFEFFc'])).toEqual(['ab', 'c']);
  });

  it('does not mutate the array it was given', () => {
    const input = ['  a  '];
    trimWhitespace(input);
    expect(input).toEqual(['  a  ']);
  });
});

describe('normalizeNulls', () => {
  it('maps every placeholder token to null', () => {
    expect(normalizeNulls(['a', '-', 'N/A', '#REF!', '']))
      .toEqual(['a', null, null, null, null]);
  });
});

describe('parseNumber', () => {
  it('coerces number-like strings and nulls the rest', () => {
    expect(parseNumber(['1,234', 'RM 5', 'nope'])).toEqual([1234, 5, null]);
  });

  it('leaves leading-zero codes alone by nulling them, never mangling them', () => {
    expect(parseNumber(['007'])).toEqual([null]);
  });
});

describe('unifyCase', () => {
  it('maps case variants onto the most frequent spelling', () => {
    expect(unifyCase(['HR', 'hr', 'HR', 'Hr'])).toEqual(['HR', 'HR', 'HR', 'HR']);
  });

  it('breaks a frequency tie by first appearance', () => {
    expect(unifyCase(['hr', 'HR'])).toEqual(['hr', 'hr']);
  });
});

describe('categoryKey', () => {
  it('ignores case, padding, internal runs and punctuation', () => {
    expect(categoryKey('  Kuala  Lumpur ')).toBe(categoryKey('kuala lumpur'));
    expect(categoryKey('PMW-SS')).toBe(categoryKey('pmw ss'));
  });
});

describe('clusterCategories', () => {
  it('groups variants under the most frequent spelling', () => {
    const clusters = clusterCategories([
      'Kuala Lumpur', 'kuala lumpur', 'Kuala  Lumpur ', 'Kuala Lumpur', 'Penang',
    ]);
    const kl = clusters.find((c) => c.canonical === 'Kuala Lumpur');
    expect(kl.count).toBe(4);
    expect(kl.variants).toHaveLength(3);
  });

  // Spec §3 and §8.2 -- these must NOT be merged.
  it('never merges genuinely different values that are merely similar', () => {
    const clusters = clusterCategories(['Dept A', 'Dept B']);
    expect(clusters).toHaveLength(2);
  });

  it('orders clusters by how many values they cover', () => {
    const clusters = clusterCategories(['a', 'b', 'b', 'c', 'c', 'c']);
    expect(clusters.map((c) => c.canonical)).toEqual(['c', 'b', 'a']);
  });
});

describe('mergeCategories', () => {
  it('rewrites values to their canonical spelling', () => {
    expect(mergeCategories(
      ['HR', 'hr', 'IT'],
      { map: { hr: 'HR' } },
    )).toEqual(['HR', 'HR', 'IT']);
  });

  it('leaves values whose key is not in the map untouched', () => {
    expect(mergeCategories(['Ops'], { map: {} })).toEqual(['Ops']);
  });
});

describe('parseDate', () => {
  it('turns dates into epoch milliseconds under the given order', () => {
    const [epoch] = parseDate(['13/01/2024'], { order: 'dmy', dateOnly: true });
    expect(new Date(epoch).toISOString()).toBe('2024-01-13T00:00:00.000Z');
  });

  it('nulls values that do not parse rather than guessing at them', () => {
    expect(parseDate(['not a date'], { order: 'dmy' })).toEqual([null]);
  });
});

describe('castType', () => {
  it('casts to numbers, nulling the casualties', () => {
    expect(castType(['1,234', 'nope'], { type: 'numeric' })).toEqual([1234, null]);
  });

  it('casts to booleans from word pairs only', () => {
    expect(castType(['Yes', 'no', '1'], { type: 'boolean' })).toEqual([true, false, null]);
  });

  it('casts to text as trimmed strings', () => {
    expect(castType([' a ', 5], { type: 'text' })).toEqual(['a', '5']);
  });
});

describe('dedupeRows', () => {
  it('removes exact duplicate rows and keeps the first', () => {
    const grid = { headers: ['a'], rows: [['x'], ['y'], ['x']] };
    expect(dedupeRows(grid).rows).toEqual([['x'], ['y']]);
  });
});

describe('dropEmptyColumns', () => {
  it('removes columns that are entirely null', () => {
    const grid = { headers: ['a', 'blank'], rows: [['x', null], ['y', null]] };
    expect(dropEmptyColumns(grid)).toEqual({ headers: ['a'], rows: [['x'], ['y']] });
  });

  it('keeps a column that has even one value', () => {
    const grid = { headers: ['a', 'sparse'], rows: [['x', null], ['y', 'z']] };
    expect(dropEmptyColumns(grid).headers).toEqual(['a', 'sparse']);
  });
});

describe('dropEmptyRows', () => {
  it('removes rows where every cell is a null token', () => {
    const grid = { headers: ['a', 'b'], rows: [['x', 'y'], ['', 'N/A'], ['z', null]] };
    expect(dropEmptyRows(grid).rows).toEqual([['x', 'y'], ['z', null]]);
  });
});
