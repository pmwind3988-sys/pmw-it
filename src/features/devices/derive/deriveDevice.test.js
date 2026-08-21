import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from './deriveDevice.js';

const load = (name) => ({
  text: readFileSync(fileURLToPath(new URL(`../__fixtures__/${name}`, import.meta.url)), 'utf8'),
  fileName: name,
  lastModified: Date.UTC(2026, 7, 19, 1, 18),
});

describe('deriveDevice — ASHRAF-PC', () => {
  const device = deriveDevice(load('ASHRAF-PC_.txt'));

  it('reads identity', () => {
    expect(device.computerName).toBe('ASHRAF-PC');
    expect(device.owner).toBe('Ashraf');
    expect(device.ownerSource).toBe('Server credential');
    expect(device.deviceType).toBe('Laptop');
    expect(device.department).toBe(null);
  });

  it('reads specs', () => {
    expect(device.installedRamGB).toBe(8);
    expect(device.storageTotalGB).toBe(477);
    expect(device.storageType).toBe('SSD only');
    expect(device.cpuGeneration).toBe('13');
    expect(device.windowsMajor).toBe(11);
  });

  it('counts real GPUs and monitors, not the virtual ones', () => {
    expect(device.gpuList).toEqual(['Intel(R) Iris(R) Xe Graphics']);
    expect(device.monitorCount).toBe(1);
  });

  it('counts mailboxes and archives separately', () => {
    expect(device.mailboxCount).toBe(1);
    expect(device.archiveCount).toBe(2);
  });

  it('scores it Watch on its 8 GB alone', () => {
    expect(device.riskScore).toBe(15);
    expect(device.riskLevel).toBe('Watch');
  });

  it('keeps the raw report for later re-derivation', () => {
    expect(device.rawReport).toContain('KBG50ZNV512G KIOXIA');
  });

  it('carries both timestamps through', () => {
    expect(device.scannedOn).toBe(Date.UTC(2026, 7, 19, 1, 18));
    expect(typeof device.importedOn).toBe('number');
    expect(device.sourceFileName).toBe('ASHRAF-PC_.txt');
  });
});

describe('deriveDevice — the awkward machines', () => {
  it('marks the failed CARMEN-HP scan incomplete with no risk score', () => {
    const device = deriveDevice(load('CARMEN-HP_.txt'));
    expect(device.scanComplete).toBe(false);
    expect(device.riskScore).toBe(null);
    expect(device.riskLevel).toBe('Unknown');
    expect(device.computerName).toBe('CARMEN-HP');
  });

  it('scores DESKTOP-8SBR420 Critical', () => {
    const device = deriveDevice(load('[STOCKYARDF1] DESKTOP-8SBR420_.txt'));
    expect(device.department).toBe('STOCKYARDF1');
    expect(device.deviceType).toBe('Desktop');
    expect(device.installedRamGB).toBe(2);
    expect(device.storageType).toBe('Mixed');
    expect(device.riskLevel).toBe('Critical');
  });

  it('reports the RAM discrepancy on EVONNE-HP and flags the free slot', () => {
    const device = deriveDevice(load('[FINANCE] EVONNE-HP_.txt'));
    expect(device.installedRamGB).toBe(8);
    expect(device.ramSlotsUsed).toBe(1);
    expect(device.ramSlotsTotal).toBe(2);
    expect(device.ramUpgradable).toBe(true);
  });

  it('collapses the 22 duplicated antivirus entries on AMIR-HP', () => {
    const device = deriveDevice(load('[ENGINEERING] AMIR-HP_.txt'));
    expect(device.antivirusProducts).toHaveLength(3);
    expect(device.avProtected).toBe(true);
    expect(device.hasHdd).toBe(true);
    expect(device.riskLevel).toBe('High');
  });

  it('reads Core Ultra on PMWL034', () => {
    const device = deriveDevice(load('PMWL034_.txt'));
    expect(device.cpuGeneration).toBe('Ultra 1');
    expect(device.cpuAgeBand).toBe('Current');
    // Current hardware throughout: its only charge is the missing managed
    // antivirus, which Defender being enabled does not excuse.
    expect(device.riskReasons).toEqual(['Managed antivirus not installed or deactivated']);
    expect(device.riskLevel).toBe('Watch');
  });

  it('reads the person out of a combined department bracket', () => {
    const device = deriveDevice(load('[QAQC FAIRUS]HPFL05_.txt'));
    expect(device.department).toBe('QAQC');
    expect(device.owner).toBe('Fairus');
    expect(device.riskLevel).toBe('Critical');
  });

  it('counts the mapped drives on PGCHAN-HP', () => {
    const device = deriveDevice(load('[SALES] PGCHAN-HP_.txt'));
    expect(device.mappedDrives).toBe(12);
    expect(device.archiveCount).toBe(6);
  });
});
