import { describe, it, expect } from 'vitest';
import { officeLicense } from './officeLicense.js';

describe('officeLicense', () => {
  it('accepts a Microsoft 365 Business install', () => {
    expect(officeLicense(['O365BusinessRetail']).licenseStatus).toBe('Authentic');
  });

  it('accepts a volume-licensed perpetual Office', () => {
    expect(officeLicense(['ProPlus2021Volume']).licenseStatus).toBe('Authentic');
  });

  it('still passes a machine that has a personal copy beside the company one', () => {
    const result = officeLicense(['O365BusinessRetail', 'O365HomePremRetail']);
    expect(result.licenseStatus).toBe('Authentic');
  });

  it('flags a personal subscription standing in for the company one', () => {
    const result = officeLicense(['O365HomePremRetail']);
    expect(result.licenseStatus).toBe('Unlicensed');
    expect(result.licenseNote).toMatch(/personal Office/);
  });

  it('flags a retail perpetual copy the company cannot account for', () => {
    expect(officeLicense(['Standard2019Retail']).licenseStatus).toBe('Unlicensed');
  });

  it('reports nothing found as undefined rather than as a failure', () => {
    expect(officeLicense([]).licenseStatus).toBe('Undefined');
  });

  it('does not treat free OneNote as an Office licence either way', () => {
    const result = officeLicense(['OneNoteFreeRetail']);
    expect(result.licenseStatus).toBe('Undefined');
    expect(result.licenseNote).toMatch(/free Microsoft apps/);
  });
});
