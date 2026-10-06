import { describe, it, expect } from 'vitest';
import { regrade } from './regrade.js';
import { defaultStandard } from '../standards/defaultStandard.js';

describe('regrade', () => {
  it('lays the grades over every device without losing its own fields', () => {
    const [d] = regrade([{ id: 3, computerName: 'PC1', department: 'HR', installedRamGB: 4, scanComplete: true }], defaultStandard());
    expect(d).toMatchObject({ id: 3, computerName: 'PC1', gradeRam: 'Critical' });
  });

  it('changes a grade when the standard changes', () => {
    const strict = defaultStandard();
    strict.profiles.desk.ram = { criticalBelow: 16, attentionBelow: 16, optimalFrom: 32 };
    expect(regrade([{ department: 'HR', installedRamGB: 8, scanComplete: true }], strict)[0].gradeRam).toBe('Critical');
  });
});
