import { describe, it, expect } from 'vitest';
import { planImport } from './syncDevices.js';

const device = (overrides) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 11 Pro', osSupported: true,
  cpuModel: 'i5', cpuAgeBand: 'Current', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 2,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 7, 19, 1, 18), sourceFileName: 'PC1_.txt',
  ...overrides,
});

describe('planImport', () => {
  it('inserts a machine the list has never seen', () => {
    const plan = planImport([device()], []);
    expect(plan.inserts).toHaveLength(1);
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('does nothing for a machine whose tracked fields are unchanged', () => {
    const existing = { ...device(), id: 7 };
    const plan = planImport([device()], [existing]);
    expect(plan.inserts).toHaveLength(0);
    expect(plan.updates).toHaveLength(0);
  });

  it('updates a machine whose RAM grew, and logs one change row', () => {
    const existing = { ...device(), id: 7 };
    const plan = planImport([device({ installedRamGB: 16 })], [existing]);

    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].id).toBe(7);
    expect(plan.changeRows).toEqual([
      {
        computerName: 'PC1', deviceId: 7, fieldName: 'installedRamGB',
        oldValue: '8', newValue: '16', changeType: 'Updated',
      },
    ]);
  });

  it('matches an existing machine case-insensitively', () => {
    const existing = { ...device({ computerName: 'pc1' }), id: 7 };
    const plan = planImport(
      [device({ computerName: 'PC1', installedRamGB: 16 })],
      [existing],
    );
    expect(plan.updates).toHaveLength(1);
  });

  it('does not update on an untracked change alone', () => {
    const existing = { ...device(), id: 7, ipAddress: '192.168.1.5' };
    const plan = planImport([device({ ipAddress: '192.168.1.99' })], [existing]);
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('carries the item body on both inserts and updates', () => {
    const plan = planImport([device()], []);
    expect(plan.inserts[0].body.Title).toBe('PC1');
    expect(plan.inserts[0].computerName).toBe('PC1');
  });

  it('counts a new-and-changed batch correctly', () => {
    const existing = { ...device({ computerName: 'PC1' }), id: 7 };
    const plan = planImport(
      [device({ installedRamGB: 16 }), device({ computerName: 'PC2' })],
      [existing],
    );
    expect(plan.inserts.map((i) => i.computerName)).toEqual(['PC2']);
    expect(plan.updates.map((u) => u.computerName)).toEqual(['PC1']);
  });
});

describe('planImport — fields set by hand', () => {
  const existing = (overrides) => ({
    ...device(), id: 7, ...overrides,
  });

  it('leaves a hand-set field alone when the file disagrees', () => {
    // Somebody corrected the owner in the register. Re-importing the same
    // unchanged scan file must not quietly undo that.
    const plan = planImport(
      [device({ owner: 'Ali' })],
      [existing({ owner: 'Ali Bin Hassan', manualFields: ['owner'] })],
    );

    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('still writes the hand-set value when something else changed', () => {
    const plan = planImport(
      [device({ owner: 'Ali', installedRamGB: 16 })],
      [existing({ owner: 'Ali Bin Hassan', manualFields: ['owner'] })],
    );

    expect(plan.updates).toHaveLength(1);
    // The body carries the manual value, not the derived one -- otherwise the
    // RAM update would overwrite the owner as a side effect.
    expect(plan.updates[0].body.Owner).toBe('Ali Bin Hassan');
    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['installedRamGB']);
  });

  it('protects only the fields named, not the whole row', () => {
    const plan = planImport(
      [device({ owner: 'Ali', department: 'ENGINEERING' })],
      [existing({
        owner: 'Ali Bin Hassan', department: 'SALES', manualFields: ['owner'],
      })],
    );

    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['department']);
    expect(plan.updates[0].body.Owner).toBe('Ali Bin Hassan');
    expect(plan.updates[0].body.Department).toBe('ENGINEERING');
  });

  it('protects several fields at once', () => {
    const plan = planImport(
      [device({ owner: 'Ali', department: 'ENGINEERING', deviceType: 'Desktop' })],
      [existing({
        owner: 'Ali Bin Hassan',
        department: 'SALES',
        deviceType: 'Laptop',
        manualFields: ['owner', 'department', 'deviceType'],
      })],
    );

    expect(plan.updates).toHaveLength(0);
  });

  it('carries the manual list forward so it is not wiped by the update', () => {
    const plan = planImport(
      [device({ installedRamGB: 16 })],
      [existing({ manualFields: ['owner'] })],
    );

    expect(plan.updates[0].body.ManualFields).toBe('owner');
  });

  it('behaves normally when nothing is hand-set', () => {
    const plan = planImport(
      [device({ owner: 'Ali' })],
      [existing({ owner: 'Ali Bin Hassan' })],
    );

    expect(plan.changeRows.map((c) => c.fieldName)).toEqual(['owner']);
    expect(plan.updates[0].body.Owner).toBe('Ali');
  });

  it('ignores a manual entry naming a field that no longer exists', () => {
    const plan = planImport(
      [device()],
      [existing({ manualFields: ['owner', 'somethingRemoved'] })],
    );

    expect(plan.updates).toHaveLength(0);
  });
});

describe('planImport — lifecycle', () => {
  const NOW = Date.UTC(2026, 9, 5);
  const scan = (over) => device({ serialNumber: 'S1', location: 'F1', ...over });
  const row = (over) => ({ ...device({ serialNumber: 'S1', location: 'F1' }), id: 7, createdOn: Date.UTC(2026, 7, 21), ...over });

  it('marks a new machine In use and opens its first stint, since at least the scan', () => {
    const plan = planImport([scan()], [], { now: NOW });
    expect(plan.inserts[0].body.Status).toBe('In use');
    expect(plan.stintOps).toHaveLength(1);
    expect(plan.stintOps[0]).toMatchObject({ computerName: 'PC1', deviceId: null, close: null });
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Ali', location: 'F1', assignedOnApprox: true });
  });

  it('renames a machine found by serial and logs the rename', () => {
    const plan = planImport([scan({ computerName: 'PC9' })], [row()], { now: NOW });
    expect(plan.updates[0]).toMatchObject({ id: 7 });
    expect(plan.updates[0].body.Title).toBe('PC9');
    expect(plan.changeRows).toContainEqual(expect.objectContaining({ fieldName: 'computerName', oldValue: 'PC1', newValue: 'PC9', deviceId: 7 }));
  });

  it('never clears a location or serial an older report does not carry', () => {
    const plan = planImport([device({ installedRamGB: 16 })], [row()], { now: NOW });
    expect(plan.updates[0].body.Location).toBe('F1');
    expect(plan.updates[0].body.SerialNumber).toBe('S1');
  });

  it('records a new owner the scan reveals as a reassignment', () => {
    const plan = planImport([scan({ owner: 'Aisyah' })], [row()], { now: NOW });
    expect(plan.stintOps).toHaveLength(1);
    expect(plan.stintOps[0].close).toMatchObject({ endReason: 'Reassigned' });
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Aisyah', assignedOnApprox: false });
  });

  it('leaves a hand-typed owner and its history alone', () => {
    const plan = planImport([scan({ owner: 'Aisyah' })], [row({ manualFields: ['owner'] })], { now: NOW });
    expect(plan.stintOps).toHaveLength(0);
  });

  it('brings a retired machine back when it is scanned again', () => {
    const plan = planImport([scan()], [row({ status: 'Retired', owner: null })], { now: NOW });
    expect(plan.updates[0].body.Status).toBe('In use');
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Ali' });
  });

  it('sends the old machine to the stash when the answer says so', () => {
    const old = row({ id: 8, computerName: 'ALI-OLD', serialNumber: 'S0' });
    const incoming = scan({ computerName: 'ALI-NEW', serialNumber: 'S2', sourceFileName: 'new.txt' });
    const plan = planImport([incoming], [old], { now: NOW, answers: { 'new.txt|8': 'stash' } });
    expect(plan.retirements).toHaveLength(1);
    expect(plan.retirements[0].body).toMatchObject({ Status: 'Spare', Owner: '' });
    expect(plan.stintOps.find((op) => op.deviceId === 8).close.endReason).toBe('Replaced');
    // The new machine is a real hand-over, not a guess about the past.
    expect(plan.stintOps.find((op) => op.computerName === 'ALI-NEW').open.assignedOnApprox).toBe(false);
  });

  it('keeps both when told to', () => {
    const old = row({ id: 8, computerName: 'ALI-OLD', serialNumber: 'S0' });
    const plan = planImport([scan({ computerName: 'ALI-NEW', serialNumber: 'S2', sourceFileName: 'new.txt' })], [old], { answers: { 'new.txt|8': 'keep' } });
    expect(plan.retirements).toHaveLength(0);
  });

  it('keeps the status and owner of a stashed machine when an older report is dropped again', () => {
    const movedOn = Date.UTC(2026, 9, 1);
    const stashed = row({ status: 'Spare', owner: '', department: null, statusChangedOn: movedOn });
    const plan = planImport([scan({ scannedOn: movedOn - 1000, installedRamGB: 16 })], [stashed], { now: NOW });
    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].body.Status).toBe('Spare');
    expect(plan.updates[0].body.Owner).toBe('');
    expect(plan.stintOps).toHaveLength(0);
  });

  it('still brings a stashed machine back when the report is newer than the move', () => {
    const movedOn = Date.UTC(2026, 9, 1);
    const stashed = row({ status: 'Spare', owner: '', statusChangedOn: movedOn });
    const plan = planImport([scan({ scannedOn: movedOn + 1000 })], [stashed], { now: NOW });
    expect(plan.updates[0].body.Status).toBe('In use');
    expect(plan.stintOps).toHaveLength(1);
  });

  it('does not stash a machine that is re-scanned in the same batch', () => {
    const old = row({ id: 8, computerName: 'ALI-OLD', serialNumber: 'S0' });
    const incoming = [
      scan({ computerName: 'ALI-OLD', serialNumber: 'S0', owner: 'Bob', sourceFileName: 'old.txt' }),
      scan({ computerName: 'ALI-NEW', serialNumber: 'S2', owner: 'Ali', sourceFileName: 'new.txt' }),
    ];
    const plan = planImport(incoming, [old], { now: NOW, answers: { 'new.txt|8': 'stash' } });
    expect(plan.prompts).toHaveLength(0);
    expect(plan.retirements).toHaveLength(0);
  });

  it('skips the second of two reports with one serial', () => {
    const plan = planImport([scan(), scan({ computerName: 'PC2', sourceFileName: 'x.txt' })], []);
    expect(plan.inserts).toHaveLength(1);
    expect(plan.skipped[0].computerName).toBe('PC2');
  });
});
