import { describe, it, expect } from 'vitest';
import { riskScore } from './riskScore.js';

const healthy = {
  osSupported: true, antivirusStatus: 'Active', avProtected: true,
  installedRamGB: 16, cpuAgeBand: 'Current', hasHdd: false, scanComplete: true,
};
const device = (overrides) => riskScore({ ...healthy, ...overrides });

describe('riskScore — individual signals', () => {
  it('scores a healthy machine zero', () => {
    expect(device({}).riskScore).toBe(0);
    expect(device({}).riskLevel).toBe('OK');
  });

  it('charges 40 for an unsupported OS', () => {
    expect(device({ osSupported: false }).riskScore).toBe(40);
  });

  it('charges 30 for missing antivirus', () => {
    expect(device({ antivirusStatus: 'Not Installed', avProtected: false }).riskScore).toBe(30);
  });

  it('charges 30 when a product is installed but nothing is enabled', () => {
    expect(device({ antivirusStatus: 'Installed — Inactive', avProtected: false }).riskScore)
      .toBe(30);
  });

  it('charges 15 for 8 GB and 25 for 4 GB, never both', () => {
    expect(device({ installedRamGB: 8 }).riskScore).toBe(15);
    expect(device({ installedRamGB: 4 }).riskScore).toBe(25);
    expect(device({ installedRamGB: 2 }).riskScore).toBe(25);
  });

  it('charges 25 for an obsolete CPU and 10 for an aging one', () => {
    expect(device({ cpuAgeBand: 'Obsolete' }).riskScore).toBe(25);
    expect(device({ cpuAgeBand: 'Aging' }).riskScore).toBe(10);
  });

  it('charges 10 for a mechanical disk', () => {
    expect(device({ hasHdd: true }).riskScore).toBe(10);
  });
});

describe('riskScore — bands', () => {
  it('places the boundaries at 20, 40 and 60', () => {
    expect(device({ hasHdd: true }).riskLevel).toBe('OK');
    expect(device({ cpuAgeBand: 'Obsolete' }).riskLevel).toBe('Watch');
    expect(device({ osSupported: false }).riskLevel).toBe('High');
    expect(device({ osSupported: false, cpuAgeBand: 'Obsolete' }).riskLevel).toBe('Critical');
  });
});

describe('riskScore — the real machines', () => {
  it('scores DESKTOP-8SBR420 at 100 — Windows 10, Pentium, 2 GB, spinning disk', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 2, cpuAgeBand: 'Obsolete', hasHdd: true, scanComplete: true,
    });
    expect(result.riskScore).toBe(100);
    expect(result.riskLevel).toBe('Critical');
  });

  it('scores HPFL05 at 80 — Windows 10, 3rd gen DDR3, 8 GB', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 8, cpuAgeBand: 'Obsolete', hasHdd: false, scanComplete: true,
    });
    expect(result.riskScore).toBe(80);
  });

  it('scores AMIR-HP at 50 — Windows 10 plus a spinning disk', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 16, cpuAgeBand: 'Current', hasHdd: true, scanComplete: true,
    });
    expect(result.riskScore).toBe(50);
  });

  it('scores ASHRAF-PC at 15 — only its 8 GB counts against it', () => {
    expect(device({ installedRamGB: 8 }).riskScore).toBe(15);
  });
});

describe('riskScore — reasons and unknowns', () => {
  it('lists a reason for every charged signal', () => {
    const result = device({ osSupported: false, hasHdd: true });
    expect(result.riskReasons).toEqual([
      'Windows 10 or older — no security updates since 14 Oct 2025',
      'Mechanical hard disk',
    ]);
  });

  it('returns a null score for an incomplete scan rather than calling it healthy', () => {
    const result = riskScore({ ...healthy, scanComplete: false });
    expect(result.riskScore).toBe(null);
    expect(result.riskLevel).toBe('Unknown');
    expect(result.riskReasons).toEqual(['Scan incomplete — re-run the report']);
  });

  it('does not charge for a signal it cannot read', () => {
    const result = riskScore({
      osSupported: null, antivirusStatus: 'Unknown', avProtected: false,
      installedRamGB: null, cpuAgeBand: 'Unknown', hasHdd: false, scanComplete: true,
    });
    expect(result.riskScore).toBe(0);
  });
});
