import { describe, it, expect } from 'vitest';
import { toneForField, toneForEntry, hasEntryTones } from './fieldTone.js';

describe('toneForField', () => {
  it('calls out the four signals the risk score charges for', () => {
    expect(toneForField({ osSupported: false }, 'osSupported')).toBe('risk');
    expect(toneForField({ antivirusStatus: 'Not Installed' }, 'antivirusStatus')).toBe('risk');
    expect(toneForField({ installedRamGB: 8 }, 'installedRamGB')).toBe('risk');
    expect(toneForField({ cpuAgeBand: 'Obsolete' }, 'cpuAgeBand')).toBe('risk');
  });

  it('calls the same signals healthy when they are', () => {
    expect(toneForField({ osSupported: true }, 'osSupported')).toBe('ok');
    expect(toneForField({ antivirusStatus: 'Active' }, 'antivirusStatus')).toBe('ok');
    expect(toneForField({ installedRamGB: 16 }, 'installedRamGB')).toBe('ok');
    expect(toneForField({ cpuAgeBand: 'Current' }, 'cpuAgeBand')).toBe('ok');
  });

  it('carries the CPU verdict onto every field that states the CPU', () => {
    const device = { cpuAgeBand: 'Aging', cpuGeneration: 'Ryzen 3000 (Zen+)' };
    expect(toneForField(device, 'cpuModel')).toBe('risk');
    expect(toneForField(device, 'cpuGeneration')).toBe('risk');
    expect(toneForField(device, 'cpuGenerationRank')).toBe('risk');
  });

  it('says nothing where there is no right answer', () => {
    expect(toneForField({ antivirusStatus: 'Unknown' }, 'antivirusStatus')).toBe(null);
    expect(toneForField({ cpuAgeBand: 'Unknown' }, 'cpuModel')).toBe(null);
    // 12 GB is neither the 8 GB the score charges for nor a machine to stop
    // thinking about.
    expect(toneForField({ installedRamGB: 12 }, 'installedRamGB')).toBe(null);
    expect(toneForField({ ipAssignment: 'Static' }, 'ipAssignment')).toBe(null);
    expect(toneForField({ ramUpgradable: true }, 'ramUpgradable')).toBe(null);
  });

  it('follows the risk score over its own Watch line', () => {
    expect(toneForField({ riskScore: 40 }, 'riskScore')).toBe('risk');
    expect(toneForField({ riskScore: 10 }, 'riskScore')).toBe('ok');
    expect(toneForField({ riskLevel: 'OK' }, 'riskLevel')).toBe('ok');
    expect(toneForField({ riskLevel: 'Watch' }, 'riskLevel')).toBe('risk');
    expect(toneForField({ riskLevel: 'Unknown' }, 'riskLevel')).toBe(null);
  });

  it('reads a mixed machine as one to look at, an all-SSD one as fine', () => {
    expect(toneForField({ storageType: 'Mixed' }, 'storageType')).toBe('risk');
    expect(toneForField({ storageType: 'SSD only' }, 'storageType')).toBe('ok');
    expect(toneForField({ storageType: 'Unknown' }, 'storageType')).toBe(null);
  });

  it('has no opinion about a device it was not given', () => {
    expect(toneForField(null, 'riskLevel')).toBe(null);
  });
});

describe('toneForEntry', () => {
  it('tones one antivirus product at a time', () => {
    expect(toneForEntry('antivirusProducts', 'Windows Defender | Enabled')).toBe('ok');
    expect(toneForEntry('antivirusProducts', 'Norton Security | Disabled')).toBe('risk');
  });

  it('tones one drive at a time', () => {
    expect(toneForEntry('storageDrivesRaw', 'KIOXIA KBG50ZNV512G | SSD | 477 GB')).toBe('ok');
    expect(toneForEntry('storageDrivesRaw', 'ST1000LM035 | Unspecified | 932 GB')).toBe('risk');
  });

  it('leaves the IT extraction disk uncoloured, since it is nobody\'s problem', () => {
    expect(toneForEntry('storageDrivesRaw', 'WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB'))
      .toBe(null);
  });

  it('knows which fields are toned entry by entry', () => {
    expect(hasEntryTones('storageDrivesRaw')).toBe(true);
    expect(hasEntryTones('gpuList')).toBe(false);
  });
});
