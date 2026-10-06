import { describe, it, expect } from 'vitest';
import { previewChanges } from './previewChanges.js';
import { defaultStandard } from './defaultStandard.js';

const device = (ram) => ({
  department: 'ENGINEERING', scanComplete: true, installedRamGB: ram, cpuGenerationRank: 12,
  storageTotalGB: 512, storageType: 'SSD only', dedicatedGpu: true, osSupported: true, windowsMajor: 11,
});

describe('previewChanges', () => {
  it('counts the machines that would change grade, per part', () => {
    const after = defaultStandard();
    after.profiles.heavy.ram.optimalFrom = 24;
    expect(previewChanges([device(24), device(28), device(16)], defaultStandard(), after))
      .toEqual([{ part: 'ram', from: 'Moderate', to: 'Optimal', count: 2 }]);
  });

  it('is empty when the standard did not change', () => {
    expect(previewChanges([device(8)], defaultStandard(), defaultStandard())).toEqual([]);
  });
});
