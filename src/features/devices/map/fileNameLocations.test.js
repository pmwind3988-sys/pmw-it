import { describe, it, expect } from 'vitest';
import { fileNameLocations } from './fileNameLocations.js';

describe('fileNameLocations', () => {
  it('leaves out one-character codes and department names a hand edit typed', () => {
    const devices = [{ location: 'E' }, { location: 'admin' }, { location: 'KL2' }];
    expect(fileNameLocations(devices)).toEqual(['F1', 'F3', 'PML', 'KL2']);
  });
});
