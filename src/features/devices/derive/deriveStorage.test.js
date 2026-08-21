import { describe, it, expect } from 'vitest';
import { deriveStorage } from './deriveStorage.js';

describe('deriveStorage', () => {
  it('classifies a single SSD machine', () => {
    expect(deriveStorage(['KBG50ZNV512G KIOXIA | SSD | 477 GB'])).toEqual({
      storageTotalGB: 477, driveCount: 1, hasHdd: false, storageType: 'SSD only',
    });
  });

  it('sums both drives and calls an SSD plus a spinning disk Mixed', () => {
    const result = deriveStorage([
      'WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB',
      'SAMSUNG MZVLQ512HBLU-00BH1 | SSD | 477 GB',
    ]);
    expect(result.storageTotalGB).toBe(1409);
    expect(result.driveCount).toBe(2);
    expect(result.hasHdd).toBe(true);
    expect(result.storageType).toBe('Mixed');
  });

  it('calls a machine with only spinning disks HDD only', () => {
    expect(deriveStorage(['WDC WD10 | Unspecified | 932 GB']).storageType).toBe('HDD only');
  });

  it('returns Unknown for a report with no storage block', () => {
    expect(deriveStorage([])).toEqual({
      storageTotalGB: null, driveCount: 0, hasHdd: false, storageType: 'Unknown',
    });
  });
});
