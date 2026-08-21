import { describe, it, expect } from 'vitest';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, TRACKED_FIELDS, toListItem, fromListItem,
} from './deviceSchema.js';

const device = {
  computerName: 'ASHRAF-PC', owner: 'Ashraf', ownerSource: 'Server credential',
  department: null, deviceType: 'Laptop', computerModel: 'HP Laptop 15-fd0xxx',
  installedRamGB: 8, reportedRamGB: 8, ramDiscrepancy: false, ramUpgradable: false,
  storageTotalGB: 477, storageType: 'SSD only', hasHdd: false,
  gpuList: ['Intel(R) Iris(R) Xe Graphics'],
  riskReasons: ['8 GB of RAM or less'], riskScore: 15, riskLevel: 'Watch',
  antivirusProducts: [{ product: 'Norton 360', enabled: true }],
  scanComplete: true,
  scannedOn: Date.UTC(2026, 7, 19, 1, 18),
  importedOn: Date.UTC(2026, 7, 21, 2, 0),
  sourceFileName: 'ASHRAF-PC_.txt', rawReport: 'Name:\n', unknownLabels: [],
};

describe('DEVICE_COLUMNS', () => {
  it('has no duplicate static names', () => {
    const names = DEVICE_COLUMNS.map((c) => c.StaticName);
    expect(new Set(names).size).toBe(names.length);
  });

  it('never declares Title as a column to create — it is built in', () => {
    expect(DEVICE_COLUMNS.some((c) => c.StaticName === 'Title')).toBe(false);
  });

  it('gives every choice column its choices', () => {
    for (const column of [...DEVICE_COLUMNS, ...CHANGE_COLUMNS]) {
      if (column.kind === 'choice') expect(column.choices?.length).toBeGreaterThan(0);
    }
  });
});

describe('toListItem', () => {
  const item = toListItem(device);

  it('puts the computer name in Title', () => {
    expect(item.Title).toBe('ASHRAF-PC');
  });

  it('sends dates as ISO instants', () => {
    expect(item.ScannedOn).toBe('2026-08-19T01:18:00.000Z');
  });

  it('mirrors the scan date as Malaysia time with AM/PM', () => {
    expect(item.ScannedOnMYT).toBe('19/08/2026 09:18 AM');
  });

  it('joins arrays with newlines', () => {
    expect(item.GpuList).toBe('Intel(R) Iris(R) Xe Graphics');
    expect(item.RiskReasons).toBe('8 GB of RAM or less');
  });

  it('renders antivirus products readably', () => {
    expect(item.AntivirusProducts).toBe('Norton 360 | Enabled');
  });

  it('sends empty string for a null text column', () => {
    expect(item.Department).toBe('');
  });

  it('omits a null number rather than sending null', () => {
    const sparse = toListItem({ ...device, installedRamGB: null });
    expect('InstalledRamGB' in sparse).toBe(false);
  });

  it('omits a null choice rather than sending null', () => {
    const sparse = toListItem({ ...device, riskLevel: null });
    expect('RiskLevel' in sparse).toBe(false);
  });

  it('sends false booleans, which are not the same as absent', () => {
    expect(item.HasHdd).toBe(false);
    expect(item.ScanComplete).toBe(true);
  });

  it('stores unknown labels as JSON', () => {
    const withExtra = toListItem({
      ...device, unknownLabels: [{ label: 'BitLocker Status', value: 'On' }],
    });
    expect(JSON.parse(withExtra.ExtraFields)).toEqual([
      { label: 'BitLocker Status', value: 'On' },
    ]);
  });
});

describe('fromListItem', () => {
  it('round-trips the fields the register needs', () => {
    const row = { ...toListItem(device), Id: 42 };
    const record = fromListItem(row);
    expect(record.id).toBe(42);
    expect(record.computerName).toBe('ASHRAF-PC');
    expect(record.installedRamGB).toBe(8);
    expect(record.hasHdd).toBe(false);
    expect(record.scannedOn).toBe(Date.UTC(2026, 7, 19, 1, 18));
    expect(record.gpuList).toEqual(['Intel(R) Iris(R) Xe Graphics']);
  });

  it('turns an absent date back into null, not NaN', () => {
    expect(fromListItem({ Title: 'X' }).scannedOn).toBe(null);
  });
});

describe('TRACKED_FIELDS', () => {
  it('tracks the hardware and health fields', () => {
    expect(TRACKED_FIELDS).toContain('installedRamGB');
    expect(TRACKED_FIELDS).toContain('riskLevel');
    expect(TRACKED_FIELDS).toContain('antivirusStatus');
  });

  it('does not track the DHCP-volatile fields', () => {
    expect(TRACKED_FIELDS).not.toContain('ipAddress');
    expect(TRACKED_FIELDS).not.toContain('ssid');
    expect(TRACKED_FIELDS).not.toContain('mappedDrives');
  });
});
