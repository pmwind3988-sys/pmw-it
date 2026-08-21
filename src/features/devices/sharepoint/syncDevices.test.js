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

describe('planSync — fields set by hand', () => {
  const existing = (overrides) => ({
    ...device(), id: 7, ...overrides,
  });

  it('leaves a hand-set field alone when the file disagrees', () => {
    // Somebody corrected the owner in the register. Re-importing the same
    // unchanged scan file must not quietly undo that.
    const plan = planSync(
      [device({ owner: 'Ali' })],
      indexByName([existing({ owner: 'Ali Bin Hassan', manualFields: ['owner'] })]),
    );

    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('still writes the hand-set value when something else changed', () => {
    const plan = planSync(
      [device({ owner: 'Ali', installedRamGB: 16 })],
      indexByName([existing({ owner: 'Ali Bin Hassan', manualFields: ['owner'] })]),
    );

    expect(plan.updates).toHaveLength(1);
    // The body carries the manual value, not the derived one -- otherwise the
    // RAM update would overwrite the owner as a side effect.
    expect(plan.updates[0].body.Owner).toBe('Ali Bin Hassan');
    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['installedRamGB']);
  });

  it('protects only the fields named, not the whole row', () => {
    const plan = planSync(
      [device({ owner: 'Ali', department: 'ENGINEERING' })],
      indexByName([existing({
        owner: 'Ali Bin Hassan', department: 'SALES', manualFields: ['owner'],
      })]),
    );

    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['department']);
    expect(plan.updates[0].body.Owner).toBe('Ali Bin Hassan');
    expect(plan.updates[0].body.Department).toBe('ENGINEERING');
  });

  it('protects several fields at once', () => {
    const plan = planSync(
      [device({ owner: 'Ali', department: 'ENGINEERING', deviceType: 'Desktop' })],
      indexByName([existing({
        owner: 'Ali Bin Hassan',
        department: 'SALES',
        deviceType: 'Laptop',
        manualFields: ['owner', 'department', 'deviceType'],
      })]),
    );

    expect(plan.updates).toHaveLength(0);
  });

  it('carries the manual list forward so it is not wiped by the update', () => {
    const plan = planSync(
      [device({ installedRamGB: 16 })],
      indexByName([existing({ manualFields: ['owner'] })]),
    );

    expect(plan.updates[0].body.ManualFields).toBe('owner');
  });

  it('behaves normally when nothing is hand-set', () => {
    const plan = planSync(
      [device({ owner: 'Ali' })],
      indexByName([existing({ owner: 'Ali Bin Hassan' })]),
    );

    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['owner']);
    expect(plan.updates[0].body.Owner).toBe('Ali');
  });

  it('ignores a manual entry naming a field that no longer exists', () => {
    const plan = planSync(
      [device()],
      indexByName([existing({ manualFields: ['owner', 'somethingRemoved'] })]),
    );

    expect(plan.updates).toHaveLength(0);
  });
});
