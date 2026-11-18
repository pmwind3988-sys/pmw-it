import { describe, it, expect } from 'vitest';
import { artKeyFor } from './categoryArtKey.js';
import { CATEGORIES } from '../../assetKinds.js';

describe('artKeyFor', () => {
  it('gives every built-in category its own picture except Other', () => {
    const keys = CATEGORIES.map(artKeyFor);
    expect(keys.filter((k) => k === 'other')).toEqual(['other']);
    expect(artKeyFor('Docking Station')).toBe('dock');
    expect(artKeyFor('PC Part')).toBe('part');
    expect(artKeyFor('Tab')).toBe('tablet');
  });
  it('matches added categories by their words', () => {
    expect(artKeyFor('Headsets')).toBe('headset');
    expect(artKeyFor('HDMI cable')).toBe('cable');
    expect(artKeyFor('Something odd')).toBe('other');
  });
});
