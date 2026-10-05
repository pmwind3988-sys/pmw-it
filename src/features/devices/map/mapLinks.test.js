import { describe, it, expect } from 'vitest';
import { mapHref, parentHref } from './mapLinks.js';

describe('mapHref', () => {
  it('addresses every level', () => {
    expect(mapHref()).toBe('/devices?view=map');
    expect(mapHref({ location: 'F1' })).toBe('/devices?view=map&location=F1');
    expect(mapHref({ location: 'F1', department: 'HR & ADMIN' })).toBe('/devices?view=map&location=F1&department=HR+%26+ADMIN');
    expect(mapHref({ place: 'stash' })).toBe('/devices?view=map&place=stash');
  });
});

describe('parentHref', () => {
  it('goes one level out', () => {
    expect(parentHref({ location: 'F1', department: 'X' })).toBe('/devices?view=map&location=F1');
    expect(parentHref({ location: 'F1' })).toBe('/devices?view=map');
    expect(parentHref({ place: 'stash' })).toBe('/devices?view=map');
  });
});
