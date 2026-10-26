import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from '../derive/deriveDevice.js';
import {
  DEVICE_COLUMNS, TRACKED_FIELDS, V2_FIELDS, toListItem, fromListItem,
} from './deviceSchema.js';

const V2_NAMES = [
  'Manufacturer', 'ProductNumber', 'ModelFamily', 'MotherboardSerial', 'DeviceUuid',
  'BiosVersion', 'BiosDate', 'Chassis', 'AssetTag', 'MemoryModules', 'DriveSerials',
  'MonitorDetails', 'WindowsBuild', 'WindowsLicenseType', 'WindowsActivation',
  'WindowsKeyEnds', 'FirmwareKey', 'OfficeLicenseType', 'OfficePersonal',
  'OfficePersonalReason', 'OfficeSignedIn', 'OfficeAccounts', 'OfficeLicensedTo',
  'OfficeTenantId', 'OfficeLicenceValidUntil', 'OfficeSubscriptions', 'OfficeVersion',
  'OfficeChannel', 'DeviceJoin', 'TenantName', 'InstalledSoftware', 'SoftwareCount',
  'ScanVersion', 'ScanRunAsAdmin',
];

const column = (name) => DEVICE_COLUMNS.find((c) => c.StaticName === name);

describe('scan version 2 columns', () => {
  it('declares every new column once', () => {
    const names = DEVICE_COLUMNS.map((c) => c.StaticName);
    for (const name of V2_NAMES) expect(names.filter((n) => n === name)).toHaveLength(1);
  });

  it('gives every choice column its choices', () => {
    for (const name of ['WindowsLicenseType', 'WindowsActivation', 'OfficeLicenseType', 'DeviceJoin']) {
      expect(column(name).choices?.length, name).toBeGreaterThan(0);
    }
  });

  it('keeps the long lists as notes under the raw-text keys', () => {
    expect(column('MemoryModules')).toMatchObject({ kind: 'note', key: 'memoryModulesRaw' });
    expect(column('DriveSerials')).toMatchObject({ kind: 'note', key: 'driveSerialsRaw' });
    expect(column('MonitorDetails')).toMatchObject({ kind: 'note', key: 'monitorDetailsRaw' });
    expect(column('InstalledSoftware')).toMatchObject({ kind: 'note', key: 'installedSoftwareRaw' });
    expect(column('OfficeAccounts')).toMatchObject({ kind: 'note', key: 'officeAccountsRaw' });
    expect(column('OfficeSubscriptions')).toMatchObject({ kind: 'note', key: 'officeSubscriptionsRaw' });
  });

  it('keeps the scan version as text, so a version this code has not met is not refused', () => {
    expect(column('ScanVersion').kind).toBe('text');
  });

  it('makes the software count a number and office personal a boolean', () => {
    expect(column('SoftwareCount').kind).toBe('number');
    expect(column('OfficePersonal').kind).toBe('boolean');
  });

  it('lists the v2 record fields for carrying over', () => {
    expect(V2_FIELDS).toEqual(expect.arrayContaining([
      'manufacturer', 'memoryModulesRaw', 'officeSignedInLicences', 'installedSoftwareRaw',
      'scanVersion', 'scanRunAsAdmin', 'softwareCount',
    ]));
    expect(V2_FIELDS).toHaveLength(V2_NAMES.length);
  });
});

describe('TRACKED_FIELDS for version 2', () => {
  it('tracks only the state-like version 2 fields', () => {
    for (const key of ['productNumber', 'windowsLicenseType', 'windowsActivation',
      'officeLicenseType', 'officePersonal', 'deviceJoin', 'softwareCount']) {
      expect(TRACKED_FIELDS).toContain(key);
    }
    expect(TRACKED_FIELDS).not.toContain('installedSoftwareRaw');
    expect(TRACKED_FIELDS).not.toContain('windowsBuild');
  });
});

describe('toListItem and fromListItem for version 2', () => {
  const load = (path) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8');
  const full = deriveDevice({
    text: load('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt'),
    fileName: '[F1 IT] V2-SAMPLE_.txt',
    lastModified: Date.UTC(2026, 9, 26, 3, 28),
  });

  it('writes the raw text, not an array object, for the list columns', () => {
    const item = toListItem(full);
    expect(item.MemoryModules.split('\n')).toHaveLength(2);
    expect(item.InstalledSoftware.split('\n')).toHaveLength(56);
    expect(item.OfficeSignedIn).toBe('o365businessretail\no365homepremretail');
    expect(item.MemoryModules).not.toContain('[object Object]');
  });

  it('writes the licence values as choice and boolean columns', () => {
    const item = toListItem(full);
    expect(item.WindowsLicenseType).toBe('Volume (KMS)');
    expect(item.OfficePersonal).toBe(true);
    expect(item.SoftwareCount).toBe(56);
    expect(item.ScanVersion).toBe('2');
  });

  it('sends an empty string for a v2 text column a v1 report left empty', () => {
    const item = toListItem({ ...full, manufacturer: null });
    expect(item.Manufacturer).toBe('');
  });

  it('reads the stored raw text back on a record, with the same keys the fresh one has', () => {
    const row = { ...toListItem(full), Id: 9 };
    const stored = fromListItem(row);
    expect(stored.memoryModulesRaw).toBe(full.memoryModulesRaw);
    expect(stored.installedSoftwareRaw).toBe(full.installedSoftwareRaw);
    expect(stored.officeSignedInLicences).toEqual(['o365businessretail', 'o365homepremretail']);
    expect(stored.officePersonal).toBe(true);
    expect(stored.softwareCount).toBe(56);
  });

  it('reads an old row, from before the columns existed, as nulls and empty lists', () => {
    const stored = fromListItem({ Id: 3, Title: 'OLD-PC' });
    expect(stored.manufacturer).toBe(null);
    expect(stored.officeSignedInLicences).toBe(null);
    expect(stored.scanVersion).toBe(null);
  });
});
