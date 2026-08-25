import { describe, it, expect } from 'vitest';
import { enrichFit } from './enrichFit.js';

describe('enrichFit', () => {
  it('reads the licence, the graphics and the server link off one row', () => {
    const row = enrichFit({
      department: 'ENGINEERING',
      deviceType: 'Desktop',
      microsoftOffice: ['O365BusinessRetail'],
      gpuList: ['Intel(R) UHD Graphics'],
      mappedDrives: 2,
      networkType: 'Wi-Fi',
      installedRamGB: 16,
      storageType: 'SSD only',
      osSupported: true,
      cpuAgeBand: 'Current',
      cpuGenerationRank: 12,
      scanComplete: true,
    });

    expect(row.licenseStatus).toBe('Authentic');
    expect(row.dedicatedGpu).toBe(false);
    expect(row.serverDependent).toBe(true);
    expect(row.fitStatus).toBe('Critical');
    expect(row.personaLabel).toMatch(/Engineering/);
  });

  it('leaves a row it was handed nothing for alone', () => {
    expect(enrichFit(null)).toBe(null);
  });
});
