import { describe, it, expect } from 'vitest';
import { deriveStorage } from './deriveStorage.js';

describe('deriveStorage', () => {
  it('classifies a single SSD machine', () => {
    expect(deriveStorage(['KBG50ZNV512G KIOXIA | SSD | 477 GB'])).toEqual({
      storageTotalGB: 477,
      driveCount: 1,
      hasHdd: false,
      storageType: 'SSD only',
      ignoredDrives: null,
    });
  });

  it('sums both drives and calls an SSD plus a spinning disk Mixed', () => {
    const result = deriveStorage([
      'ST1000LM035-1RK172 | Unspecified | 932 GB',
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
      storageTotalGB: null,
      driveCount: 0,
      hasHdd: false,
      storageType: 'Unknown',
      ignoredDrives: null,
    });
  });
});

describe('deriveStorage — IT extraction disk', () => {
  const IT_DISK = 'WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB';

  it('leaves the disk out of the size, the count and the disk type', () => {
    const result = deriveStorage([IT_DISK, 'SAMSUNG MZVLQ512HBLU-00BH1 | SSD | 477 GB']);
    expect(result.storageTotalGB).toBe(477);
    expect(result.driveCount).toBe(1);
    expect(result.hasHdd).toBe(false);
    expect(result.storageType).toBe('SSD only');
  });

  it('reports what it left out, so the page can say why', () => {
    const result = deriveStorage([IT_DISK, 'Apacer AS340 240GB | SSD | 224 GB']);
    expect(result.ignoredDrives).toEqual(['WDC WD10 JPVX-60JC3T1 | 932 GB']);
  });

  it('still counts a genuine 1 TB Western Digital disk of another model', () => {
    const result = deriveStorage(['WDC WD10EZEX-08WN4A0 | Unspecified | 932 GB']);
    expect(result.driveCount).toBe(1);
    expect(result.storageType).toBe('HDD only');
    expect(result.ignoredDrives).toBe(null);
  });

  it('leaves a machine scanned with nothing else attached Unknown, not HDD only', () => {
    const result = deriveStorage([IT_DISK]);
    expect(result.storageType).toBe('Unknown');
    expect(result.driveCount).toBe(0);
    expect(result.hasHdd).toBe(false);
  });
});
