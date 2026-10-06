import { describe, it, expect } from 'vitest';
import { diffStandard, summaryOf } from './diffStandard.js';
import { defaultStandard } from './defaultStandard.js';

describe('diffStandard', () => {
  it('says nothing when nothing changed', () => {
    expect(summaryOf(diffStandard(defaultStandard(), defaultStandard()))).toBe('No changes');
  });

  it('reads like a sentence for every kind of change', () => {
    const after = defaultStandard();
    after.profiles.heavy.ram.optimalFrom = 24;
    after.profiles.desk.windows.win10Supported = 'Optimal';
    after.departments.QAQC = 'desk';
    after.colors.Critical = '#b91c1c';
    expect(diffStandard(defaultStandard(), after)).toEqual([
      'Engineering RAM optimal from 32 → 24',
      'Desk Windows 10 (supported) Moderate → Optimal',
      'QAQC judged as Desk',
      'Critical colour changed',
    ]);
  });
});
