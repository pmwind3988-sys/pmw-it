import { describe, it, expect } from 'vitest';
import { diffDevice, indexByName } from './diffDevice.js';

const existing = {
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 10 Pro', osSupported: false,
  cpuModel: 'i5', cpuAgeBand: 'Aging', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 1,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'High',
  ipAddress: '192.168.1.10', ssid: 'PMW_Group',
};

describe('diffDevice', () => {
  it('finds nothing when nothing changed', () => {
    expect(diffDevice(existing, { ...existing })).toEqual([]);
  });

  it('reports an upgrade as an Updated change', () => {
    const changes = diffDevice(existing, { ...existing, installedRamGB: 16 });
    expect(changes).toEqual([
      { fieldName: 'installedRamGB', oldValue: '8', newValue: '16', changeType: 'Updated' },
    ]);
  });

  it('reports a newly filled field as Added', () => {
    const changes = diffDevice({ ...existing, owner: null }, existing);
    expect(changes).toEqual([
      { fieldName: 'owner', oldValue: '', newValue: 'Ali', changeType: 'Added' },
    ]);
  });

  it('reports a cleared field as Removed', () => {
    const changes = diffDevice(existing, { ...existing, department: null });
    expect(changes).toEqual([
      { fieldName: 'department', oldValue: 'SALES', newValue: '', changeType: 'Removed' },
    ]);
  });

  it('ignores changes to DHCP-volatile fields', () => {
    const changes = diffDevice(existing, { ...existing, ipAddress: '192.168.1.99', ssid: 'Other' });
    expect(changes).toEqual([]);
  });

  it('treats a boolean flip as a change', () => {
    const changes = diffDevice(existing, { ...existing, osSupported: true });
    expect(changes).toEqual([
      { fieldName: 'osSupported', oldValue: 'false', newValue: 'true', changeType: 'Updated' },
    ]);
  });

  it('does not report a change when a number arrives as its string form', () => {
    expect(diffDevice(existing, { ...existing, installedRamGB: '8' })).toEqual([]);
  });

  it('reports several changes at once, in tracked-field order', () => {
    const changes = diffDevice(existing, {
      ...existing, installedRamGB: 16, riskLevel: 'Watch',
    });
    expect(changes.map((c) => c.fieldName)).toEqual(['installedRamGB', 'riskLevel']);
  });
});

describe('indexByName', () => {
  it('keys on a lower-cased computer name', () => {
    const index = indexByName([{ computerName: 'ASHRAF-PC' }]);
    expect(index.get('ashraf-pc')).toBeDefined();
  });

  it('skips rows with no computer name rather than colliding on empty', () => {
    const index = indexByName([{ computerName: null }, { computerName: 'PC1' }]);
    expect(index.size).toBe(1);
  });
});
