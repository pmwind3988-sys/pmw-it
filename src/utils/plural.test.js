import { describe, it, expect } from 'vitest';
import { plural } from './plural.js';

describe('plural', () => {
  it('agrees with the count', () => {
    expect(plural(1, 'desktop')).toBe('1 desktop');
    expect(plural(0, 'desktop')).toBe('0 desktops');
    expect(plural(3, 'machine')).toBe('3 machines');
    expect(plural(2, 'person', 'people')).toBe('2 people');
  });
});
