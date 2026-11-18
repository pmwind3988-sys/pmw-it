import { describe, it, expect } from 'vitest';
import { storeAction, STORE } from './storedReceipt';

const batch = (n) => ({ id: 'b1', drafts: Array.from({ length: n }, (_, i) => ({ localId: `d${i}` })) });

describe('what happens to the device copy of a receipt', () => {
  it('a receipt with lines is saved', () => {
    expect(storeAction(batch(2), null)).toBe(STORE.SAVE);
    expect(storeAction(batch(1), 'b1')).toBe(STORE.SAVE);
  });

  it('a receipt emptied after it was stored drops the stored copy', () => {
    expect(storeAction(batch(0), 'b1')).toBe(STORE.DROP);
  });

  it('a fresh empty receipt that was never stored does nothing', () => {
    expect(storeAction(batch(0), null)).toBe(STORE.NONE);
  });

  it('an empty receipt never drops some other stored batch', () => {
    expect(storeAction(batch(0), 'someone-else')).toBe(STORE.NONE);
  });
});
