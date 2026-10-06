import { describe, it, expect } from 'vitest';
import { validateStandard } from './validateStandard.js';
import { defaultStandard } from './defaultStandard.js';

const broken = (edit) => { const s = defaultStandard(); edit(s); return validateStandard(s); };

describe('validateStandard', () => {
  it('refuses cut-offs out of order, naming where', () => {
    const result = broken((s) => { s.profiles.heavy.ram.criticalBelow = 40; });
    expect(result.ok).toBe(false);
    expect(result.errors[0]).toEqual({ path: 'profiles.heavy.ram', message: 'Must run in order: Critical below ≤ Needs attention below ≤ Optimal from.' });
  });

  it('accepts equal cut-offs: an empty band', () => {
    expect(broken((s) => { s.profiles.desk.cpu = { criticalBelow: 8, attentionBelow: 8, optimalFrom: 8 }; }).ok).toBe(true);
  });

  it('refuses a missing or negative number', () => {
    expect(broken((s) => { s.profiles.desk.cpu.optimalFrom = null; }).ok).toBe(false);
    expect(broken((s) => { s.profiles.desk.cpu.criticalBelow = -1; }).ok).toBe(false);
  });

  it('refuses a choice that is not a grade', () => {
    expect(broken((s) => { s.profiles.mobile.windows.win11 = 'Great'; }).errors[0].path).toBe('profiles.mobile.windows.win11');
  });

  it('refuses a missing profile, an unknown profile on a department, and a bad colour', () => {
    expect(broken((s) => { delete s.profiles.desk; }).ok).toBe(false);
    expect(broken((s) => { s.departments.HR = 'boss'; }).errors[0].path).toBe('departments.HR');
    expect(broken((s) => { s.colors.Optimal = 'green'; }).errors[0].path).toBe('colors.Optimal');
  });

  it('refuses a standard from a newer schema, and nothing at all', () => {
    expect(broken((s) => { s.schema = 2; }).ok).toBe(false);
    expect(validateStandard(null).ok).toBe(false);
  });
});
