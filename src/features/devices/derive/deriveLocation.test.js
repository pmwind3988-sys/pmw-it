import { describe, it, expect } from 'vitest';
import { splitLocation } from './deriveLocation.js';

describe('splitLocation', () => {
  it('takes a known code off the front of the bracket', () => {
    expect(splitLocation('F1 ENGINEERING')).toEqual({ location: 'F1', rest: 'ENGINEERING' });
    expect(splitLocation('pml finance evonne')).toEqual({ location: 'PML', rest: 'finance evonne' });
  });

  it('leaves a bracket with no location exactly as it was', () => {
    expect(splitLocation('ENGINEERING')).toEqual({ location: null, rest: 'ENGINEERING' });
    expect(splitLocation('QAQC FAIRUS')).toEqual({ location: null, rest: 'QAQC FAIRUS' });
  });

  it('reads a bracket that is only a location', () => {
    expect(splitLocation('F3')).toEqual({ location: 'F3', rest: null });
  });

  it('splits a code glued to the end of the department', () => {
    expect(splitLocation('STOCKYARDF1')).toEqual({ location: 'F1', rest: 'STOCKYARD' });
  });

  it('splits the two-word guardhouse name', () => {
    expect(splitLocation('PML GUARDHOUSE')).toEqual({ location: 'PML', rest: 'GUARDHOUSE' });
  });

  it('does not cut a short word that only ends like a code', () => {
    expect(splitLocation('XF1')).toEqual({ location: null, rest: 'XF1' });
  });

  it('recognises a location added in the portal', () => {
    expect(splitLocation('HQ ADMIN', ['F1', 'HQ'])).toEqual({ location: 'HQ', rest: 'ADMIN' });
  });

  it('has nothing to split without a bracket', () => {
    expect(splitLocation(null)).toEqual({ location: null, rest: null });
  });
});
