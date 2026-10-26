import { describe, it, expect } from 'vitest';
import { carryV2Fields, planImport } from './syncDevices.js';
import { V2_FIELDS, toListItem } from './deviceSchema.js';

const v1 = (overrides = {}) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  installedRamGB: 8, storageTotalGB: 477, antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 7, 19, 1, 18), sourceFileName: 'PC1_.txt',
  scanVersion: '1', scanRunAsAdmin: null,
  ...overrides,
});

/** A stored row as read back from SharePoint after a full v2 scan. */
const storedV2 = {
  id: 7,
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  installedRamGB: 8, storageTotalGB: 477, antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 9, 26, 3, 28), sourceFileName: 'PC1_.txt',
  status: 'In use', statusChangedOn: Date.UTC(2026, 7, 1), location: 'F1', serialNumber: 'SERIAL1',
  scanVersion: '2', scanRunAsAdmin: false,
  manufacturer: 'HP', productNumber: '7Z8D7PA#UUF', memoryModulesRaw: 'Bottom - Slot 1 (left) | 4 GB',
  windowsLicenseType: 'OEM', windowsActivation: 'Activated', officeLicenseType: 'Microsoft 365',
  officePersonal: true, officeSignedInLicences: ['o365homepremretail'],
  deviceJoin: 'Azure AD', softwareCount: 56, installedSoftwareRaw: 'Google Chrome | 1 | Google |',
};

describe('carryV2Fields', () => {
  it('keeps the stored version 2 fields when a version 1 report arrives', () => {
    const merged = carryV2Fields(v1({ serialNumber: 'SERIAL1' }), storedV2);
    expect(merged.manufacturer).toBe('HP');
    expect(merged.memoryModulesRaw).toBe('Bottom - Slot 1 (left) | 4 GB');
    expect(merged.windowsLicenseType).toBe('OEM');
    expect(merged.officeSignedInLicences).toEqual(['o365homepremretail']);
    expect(merged.officePersonal).toBe(true);
    expect(merged.softwareCount).toBe(56);
    expect(merged.installedSoftwareRaw).toBe('Google Chrome | 1 | Google |');
  });

  it('keeps the stored scan version, so the carried fields still say where they came from', () => {
    expect(carryV2Fields(v1(), storedV2).scanVersion).toBe('2');
    expect(carryV2Fields(v1(), storedV2).scanRunAsAdmin).toBe(false);
  });

  it('keeps the v1 fields the new report did bring', () => {
    const merged = carryV2Fields(v1({ installedRamGB: 16 }), storedV2);
    expect(merged.installedRamGB).toBe(16);
  });

  it('keeps a stored full scan over a basic one, but not the other way round', () => {
    const basic = v1({ scanVersion: '2 basic', manufacturer: null });
    expect(carryV2Fields(basic, storedV2).manufacturer).toBe('HP');

    const fullIncoming = v1({ scanVersion: '2', manufacturer: 'Dell' });
    expect(carryV2Fields(fullIncoming, storedV2)).toBe(fullIncoming);
  });

  it('lets a basic scan bring what it did read, and keeps only what it could not', () => {
    const basic = v1({
      scanVersion: '2 basic', windowsActivation: 'Not activated', softwareCount: 49,
      installedSoftwareRaw: 'Norton 360 | 24 | Gen Digital |', memoryModulesRaw: null, deviceJoin: 'Unknown',
    });
    const merged = carryV2Fields(basic, storedV2);
    expect(merged.windowsActivation).toBe('Not activated');
    expect(merged.softwareCount).toBe(49);
    expect(merged.installedSoftwareRaw).toBe('Norton 360 | 24 | Gen Digital |');
    expect(merged.memoryModulesRaw).toBe('Bottom - Slot 1 (left) | 4 GB');
    expect(merged.deviceJoin).toBe('Azure AD');
  });

  it('writes the incoming report as it is when the stored row is also version 1 or older', () => {
    const incoming = v1({ manufacturer: 'HP' });
    expect(carryV2Fields(incoming, { ...storedV2, scanVersion: '1' })).toBe(incoming);
    expect(carryV2Fields(incoming, { ...storedV2, scanVersion: null })).toBe(incoming);
  });

  it('writes normally when there is no stored row', () => {
    const incoming = v1();
    expect(carryV2Fields(incoming, undefined)).toBe(incoming);
  });

  it('does not change the incoming object', () => {
    const incoming = v1();
    carryV2Fields(incoming, storedV2);
    expect(incoming.manufacturer).toBeUndefined();
  });

  it('carries every version 2 field the schema names', () => {
    const merged = carryV2Fields(v1(), storedV2);
    for (const key of V2_FIELDS) expect(key in merged, key).toBe(true);
  });
});

describe('planImport with a version 1 report over a version 2 row', () => {
  it('does not blank the stored version 2 text columns on the update', () => {
    const plan = planImport([v1({ installedRamGB: 16 })], [storedV2]);
    expect(plan.updates).toHaveLength(1);

    const body = plan.updates[0].body;
    expect(body.Manufacturer).toBe('HP');
    expect(body.MemoryModules).toBe('Bottom - Slot 1 (left) | 4 GB');
    expect(body.WindowsLicenseType).toBe('OEM');
    expect(body.ScanVersion).toBe('2');
  });

  it('logs no change for a v2 field a v1 report cannot speak to', () => {
    const plan = planImport([v1({ installedRamGB: 16 })], [storedV2]);
    const fields = plan.changeRows.map((row) => row.fieldName);
    expect(fields).toEqual(['installedRamGB']);
  });

  it('keeps the v2 fields of a row whose update is only the v1 part', () => {
    const plan = planImport([v1({ riskLevel: 'High' })], [storedV2]);
    expect(plan.updates[0].body.DeviceJoin).toBe('Azure AD');
    expect(toListItem({ ...storedV2 }).DeviceJoin).toBe('Azure AD');
  });
});
