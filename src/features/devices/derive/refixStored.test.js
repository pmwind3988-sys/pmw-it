import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { refixStored } from './refixStored.js';

describe('refixStored — storage summary self-correcting on read', () => {
  it('drops the IT USB stick a stale row still counts, and fixes the type', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nUSB DISK 2.0 | Unspecified | 15 GB',
      storageTotalGB: 462,
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      // No risk inputs that would score anything else, so the only charge is
      // the phantom spinning disk.
      scanComplete: true,
    };

    const fixed = refixStored(stored);

    expect(fixed.storageType).toBe('SSD only');
    expect(fixed.driveCount).toBe(1);
    expect(fixed.hasHdd).toBe(false);
    expect(fixed.storageTotalGB).toBe(447);
  });

  it('clears the spinning-disk risk charge when the only disk was the IT stick', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nUSB DISK 2.0 | Unspecified | 15 GB',
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      scanComplete: true,
      // A healthy machine but for the phantom disk: no other charge to muddy
      // the before/after.
      osSupported: true,
      avProtected: true,
      antivirusStatus: 'Active',
      installedRamGB: 16,
      cpuAgeBand: 'Current',
    };

    const fixed = refixStored(stored);

    expect(fixed.riskReasons).not.toContain('Mechanical hard disk');
    expect(fixed.riskScore).toBe(0);
  });

  it('leaves a genuine mixed machine as Mixed', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB\nWDC WD10 SPZX | Unspecified | 932 GB',
      driveCount: 2,
      hasHdd: true,
      storageType: 'Mixed',
      scanComplete: true,
    };

    const fixed = refixStored(stored);

    expect(fixed.storageType).toBe('Mixed');
    expect(fixed.hasHdd).toBe(true);
    expect(fixed.driveCount).toBe(2);
  });

  it('leaves the storage figures of a row with no raw drive list alone', () => {
    const stored = { storageType: 'Mixed', driveCount: 2, hasHdd: true };
    expect(refixStored(stored)).toMatchObject(stored);
  });

  it('keeps the stored storage figures when they already agree with the raw list', () => {
    const stored = {
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB',
      storageTotalGB: 447,
      driveCount: 1,
      hasHdd: false,
      storageType: 'SSD only',
      scanComplete: true,
    };
    expect(refixStored(stored)).toMatchObject(stored);
  });

  it('lays the persona verdict over every row it reads', () => {
    const fixed = refixStored({
      storageDrivesRaw: 'Apacer AS340 480GB | SSD | 447 GB',
      department: 'SALES',
      deviceType: 'Laptop',
      microsoftOffice: ['O365BusinessRetail'],
      scanComplete: true,
    });

    expect(fixed.personaLabel).toMatch(/Executive/);
    expect(fixed.licenseStatus).toBe('Authentic');
    expect(fixed.suggestedFormFactor).toBe('Laptop');
  });
});

describe('refixStored — a new scan imported before its sections were known', () => {
  // What the register holds for a machine whose version 2 report was imported
  // by a build that did not yet know the new labels: the report text is all
  // there, but every new block went to Extra Fields and nothing says "v2".
  const text = readFileSync(new URL('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt', import.meta.url), 'utf8');
  const stale = {
    computerName: 'V2-SAMPLE',
    scanVersion: null,
    windowsLicenseType: null,
    productNumber: null,
    memoryModulesRaw: null,
    installedSoftwareRaw: null,
    softwareCount: null,
    extraFields: JSON.stringify([{ label: 'Product Number', value: '7Z8D7PA#UUF' }]),
    rawReport: text,
  };

  it('reads the licences, part numbers and software back out of the stored report', () => {
    const fixed = refixStored(stale);

    expect(fixed.scanVersion).toBe('2');
    expect(fixed.windowsLicenseType).toBe('Volume (KMS)');
    expect(fixed.productNumber).toBe('7Z8D7PA#UUF');
    expect(fixed.memoryModulesRaw).toMatch(/S\/N 3430987E/);
    expect(fixed.installedSoftwareRaw).toMatch(/4Team Safe PST Backup/);
    expect(fixed.softwareCount).toBeGreaterThan(0);
  });

  it('stops listing the blocks it now understands as extra fields', () => {
    expect(refixStored(stale).extraFields).toBeNull();
  });

  it('leaves a row alone when its stored report really is an old scan', () => {
    const old = { ...stale, rawReport: 'Computer Name:\nPC1\n\nTotal RAM:\n8 GB\n' };
    expect(refixStored(old).scanVersion).toBeNull();
    expect(refixStored(old).extraFields).toBe(stale.extraFields);
  });

  it('rebuilds the report from Extra Fields when SharePoint cut the stored text short', () => {
    // PMWL013: the stored report stops mid-software list, so its Scan Info is
    // gone from the text, but every block the old build did not know was kept
    // whole in Extra Fields.
    const cut = text.indexOf('Manufacturer:');
    // What the old build filed as unknown: each block's label and the lines under it.
    const unknownLabels = text.slice(cut).replace(/\r\n?/g, '\n').split(/\n\s*\n/)
      .map((block) => block.trim().split('\n'))
      .filter(([first]) => /^[A-Z][\w ]+:$/.test(first))
      .map(([first, ...rest]) => ({ label: first.slice(0, -1), value: rest.join('\n') }));
    const truncated = {
      ...stale,
      rawReport: text.slice(0, text.indexOf('Installed Software:') + 40),
      extraFields: JSON.stringify(unknownLabels),
    };
    const fixed = refixStored(truncated);

    expect(fixed.scanVersion).toBe('2');
    expect(fixed.windowsLicenseType).toBe('Volume (KMS)');
    expect(fixed.memoryModulesRaw).toMatch(/S\/N 3430987E/);
    expect(fixed.softwareCount).toBe(refixStored(stale).softwareCount);
  });

  it('never re-reads a row that already holds a new scan', () => {
    const held = { ...stale, scanVersion: '2', windowsLicenseType: 'OEM' };
    expect(refixStored(held).windowsLicenseType).toBe('OEM');
  });
});
