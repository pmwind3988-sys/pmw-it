import { describe, it, expect } from 'vitest';
import { BUILT_IN_LOCATIONS, cleanLocation, locationsIn } from './locations.js';

describe('cleanLocation', () => {
  it('upper-cases and trims', () => expect(cleanLocation('  f1 ')).toBe('F1'));
  it('reads a blank as no location', () => {
    expect(cleanLocation('')).toBeNull();
    expect(cleanLocation(null)).toBeNull();
  });
});

describe('locationsIn', () => {
  it('starts with the built-in locations', () => {
    expect(locationsIn([])).toEqual(BUILT_IN_LOCATIONS);
    expect(BUILT_IN_LOCATIONS).toEqual(['F1', 'F3', 'PML']);
  });

  it('adds a location the register is using, once, whatever its case', () => {
    expect(locationsIn([{ location: 'hq' }, { location: 'HQ' }, { location: 'f1' }, { location: '' }]))
      .toEqual(['F1', 'F3', 'PML', 'HQ']);
  });
});
