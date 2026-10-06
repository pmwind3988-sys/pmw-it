import { describe, it, expect } from 'vitest';
import { createCooldown } from './cooldown.js';

describe('createCooldown', () => {
  it('takes a code held in frame once, and again only after it was away', () => {
    const fresh = createCooldown(1000);
    expect(fresh('abc', 0)).toBe(true);
    expect(fresh('ABC', 30)).toBe(false);
    expect(fresh('abc', 900)).toBe(false);
    expect(fresh('abc', 1800)).toBe(false);
    expect(fresh('abc', 3000)).toBe(true);
    expect(fresh('other', 3000)).toBe(true);
    expect(fresh('', 3000)).toBe(false);
  });
});
