import { describe, it, expect } from 'vitest';
import { planSync } from './syncDevices.js';
import { indexByName } from './diffDevice.js';

const device = (overrides) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 11 Pro', osSupported: true,
  cpuModel: 'i5', cpuAgeBand: 'Current', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 2,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 7, 19, 1, 18), sourceFileName: 'PC1_.txt',
  ...overrides,
});

describe('planSync', () => {
  it('inserts a machine the list has never seen', () => {
    const plan = planSync([device()], indexByName([]));
    expect(plan.inserts).toHaveLength(1);
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('does nothing for a machine whose tracked fields are unchanged', () => {
    const existing = { ...device(), id: 7 };
    const plan = planSync([device()], indexByName([existing]));
    expect(plan.inserts).toHaveLength(0);
    expect(plan.updates).toHaveLength(0);
  });

  it('updates a machine whose RAM grew, and logs one change row', () => {
    const existing = { ...device(), id: 7 };
    const plan = planSync([device({ installedRamGB: 16 })], indexByName([existing]));

    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].id).toBe(7);
    expect(plan.changeRows).toEqual([
      {
        computerName: 'PC1', fieldName: 'installedRamGB',
        oldValue: '8', newValue: '16', changeType: 'Updated',
      },
    ]);
  });

  it('matches an existing machine case-insensitively', () => {
    const existing = { ...device({ computerName: 'pc1' }), id: 7 };
    const plan = planSync(
      [device({ computerName: 'PC1', installedRamGB: 16 })],
      indexByName([existing]),
    );
    expect(plan.updates).toHaveLength(1);
  });

  it('does not update on an untracked change alone', () => {
    const existing = { ...device(), id: 7, ipAddress: '192.168.1.5' };
    const plan = planSync([device({ ipAddress: '192.168.1.99' })], indexByName([existing]));
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('carries the item body on both inserts and updates', () => {
    const plan = planSync([device()], indexByName([]));
    expect(plan.inserts[0].body.Title).toBe('PC1');
    expect(plan.inserts[0].computerName).toBe('PC1');
  });

  it('counts a new-and-changed batch correctly', () => {
    const existing = { ...device({ computerName: 'PC1' }), id: 7 };
    const plan = planSync(
      [device({ installedRamGB: 16 }), device({ computerName: 'PC2' })],
      indexByName([existing]),
    );
    expect(plan.inserts.map((i) => i.computerName)).toEqual(['PC2']);
    expect(plan.updates.map((u) => u.computerName)).toEqual(['PC1']);
  });
});
