import { describe, it, expect } from 'vitest';
import { typedQuantity, settledQuantity } from './quantityText.js';

describe('typedQuantity', () => {
  it('reads a number that has been typed', () => {
    expect(typedQuantity('3')).toBe(3);
    expect(typedQuantity(' 12 ')).toBe(12);
  });

  it('commits nothing for an empty box, so the row keeps its count', () => {
    // The bug this exists for: an empty box read as a number put a 1 back
    // under the cursor, and "3" came out as 13.
    expect(typedQuantity('')).toBeNull();
    expect(typedQuantity('   ')).toBeNull();
    expect(typedQuantity(null)).toBeNull();
  });

  it('commits nothing for a number a row cannot have', () => {
    expect(typedQuantity('0')).toBeNull();
    expect(typedQuantity('-4')).toBeNull();
    expect(typedQuantity('lots')).toBeNull();
  });

  it('counts whole things', () => {
    expect(typedQuantity('2.7')).toBe(2);
  });
});

describe('settledQuantity', () => {
  it('takes the typed number when there is one', () => {
    expect(settledQuantity('3', 1)).toBe(3);
  });

  it('puts the previous count back when the box was left empty', () => {
    expect(settledQuantity('', 4)).toBe(4);
    expect(settledQuantity('0', 4)).toBe(4);
  });

  it('falls back to one when there was no previous count either', () => {
    expect(settledQuantity('', undefined)).toBe(1);
    expect(settledQuantity('', 0)).toBe(1);
  });
});
