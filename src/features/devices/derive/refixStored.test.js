import { describe, it, expect } from 'vitest';
import { refixStored } from './refixStored.js';

describe('refixStored — storage summary self-correcting on read', () => {
  it('drops the IT USB stick a stale row still counts, and fixes the type', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nUSB DISK 2.0 | Unspecified | 15 GB',
      storageTotalGB: 462,
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      // No risk inputs that would score anything else, so the only charge is
      // the phantom spinning disk.
      scanComplete: true,
    };

    const fixed = refixStored(stored);

    expect(fixed.storageType).toBe('SSD only');
    expect(fixed.driveCount).toBe(1);
    expect(fixed.hasHdd).toBe(false);
    expect(fixed.storageTotalGB).toBe(447);
  });

  it('clears the spinning-disk risk charge when the only disk was the IT stick', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nUSB DISK 2.0 | Unspecified | 15 GB',
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      scanComplete: true,
      // A healthy machine but for the phantom disk: no other charge to muddy
      // the before/after.
      osSupported: true,
      avProtected: true,
      antivirusStatus: 'Active',
      installedRamGB: 16,
      cpuAgeBand: 'Current',
    };

    const fixed = refixStored(stored);

    expect(fixed.riskReasons).not.toContain('Mechanical hard disk');
    expect(fixed.riskScore).toBe(0);
  });

  it('leaves a genuine mixed machine as Mixed', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nWDC WD10 SPZX | Unspecified | 932 GB',
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      scanComplete: true,
    };

    const fixed = refixStored(stored);

    expect(fixed.storageType).toBe('Mixed');
    expect(fixed.hasHdd).toBe(true);
    expect(fixed.driveCount).toBe(2);
  });

  it('leaves a row with no raw drive list untouched', () => {
    const stored = { storageType: 'Mixed', driveCount: 2, hasHdd: true };
    expect(refixStored(stored)).toBe(stored);
  });

  it('returns the same object when nothing changed, so reads stay cheap', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB',
      storageTotalGB: 447,
      driveCount: 1,
      hasHdd: false,
      storageType: 'SSD only',
      scanComplete: true,
    };
    expect(refixStored(stored)).toBe(stored);
  });
});
