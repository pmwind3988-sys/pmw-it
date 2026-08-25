import { describe, it, expect } from 'vitest';
import { deviceFit } from './deviceFit.js';

/** A machine that clears every bar for a desk worker. */
const base = {
  department: 'FINANCE',
  deviceType: 'Desktop',
  installedRamGB: 16,
  storageType: 'SSD only',
  hasHdd: false,
  osSupported: true,
  cpuAgeBand: 'Current',
  cpuGenerationRank: 12,
  dedicatedGpu: false,
  licenseStatus: 'Authentic',
  serverDependent: false,
  networkRisk: 'None',
  scanComplete: true,
};

describe('deviceFit', () => {
  it('calls a comfortable, licensed, modern desk machine optimal', () => {
    expect(deviceFit(base).fitStatus).toBe('Optimal');
  });

  it('drops a machine to critical on an unlicensed Office', () => {
    const result = deviceFit({ ...base, licenseStatus: 'Unlicensed' });
    expect(result.fitStatus).toBe('Critical');
    expect(result.fitReasons[0]).toMatch(/outside the company licence/);
  });

  it('drops a machine to critical on a spinning boot disk', () => {
    expect(deviceFit({ ...base, storageType: 'HDD only', hasHdd: true }).fitStatus)
      .toBe('Critical');
  });

  it('holds engineering to a higher memory floor than the desk', () => {
    const thin = { ...base, installedRamGB: 8 };
    expect(deviceFit(thin).fitStatus).toBe('Moderate');
    expect(deviceFit({ ...thin, department: 'ENGINEERING' }).fitStatus).toBe('Needs Attention');
    expect(deviceFit({ ...thin, department: 'ENGINEERING' }).fitReasons[0])
      .toMatch(/under the 16 GB/);
  });

  it('flags drawing work running on processor graphics', () => {
    const result = deviceFit({ ...base, department: 'ENGINEERING', installedRamGB: 32 });
    expect(result.fitStatus).toBe('Needs Attention');
    expect(result.fitReasons).toContain('Drawing work on processor graphics — no dedicated card fitted');
  });

  it('treats server work over Wi-Fi as a critical dependency', () => {
    const result = deviceFit({
      ...base, serverDependent: true, networkRisk: 'Severe',
    });
    expect(result.fitStatus).toBe('Critical');
    expect(result.actionRequired).toMatch(/bottleneck/);
  });

  it('suggests a laptop for a field role without calling the desktop a fault', () => {
    const result = deviceFit({ ...base, department: 'SALES' });
    expect(result.suggestedFormFactor).toBe('Laptop');
    expect(result.fitStatus).not.toBe('Critical');
    expect(result.fitReasons.join(' ')).not.toMatch(/laptop/i);
  });

  it('does not call a machine optimal while its form factor is the wrong one', () => {
    const result = deviceFit({ ...base, department: 'SALES', installedRamGB: 16 });
    expect(result.formFactorMatches).toBe(false);
    expect(result.fitStatus).toBe('Moderate');
  });

  it('judges an unclassified machine against the desk baseline', () => {
    const result = deviceFit({ ...base, department: null });
    expect(result.personaLabel).toBe('Unclassified');
    expect(result.suggestedFormFactor).toBe(null);
  });

  it('refuses to grade a machine whose scan failed', () => {
    const result = deviceFit({ ...base, scanComplete: false });
    expect(result.fitStatus).toBe('Unknown');
    expect(result.actionRequired).toBe('Re-run the scan');
  });
});
