import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from './deriveDevice.js';
import { parseReport } from '../parse/parseReport.js';
import { hardwareIdsOf } from './deriveHardwareIds.js';
import { officeAccountsOf } from './deriveLicensing.js';
import { softwareOf } from './deriveSoftware.js';

const load = (path, fileName) => ({
  text: readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8'),
  fileName,
  lastModified: Date.UTC(2026, 9, 26, 3, 28),
});

const FULL = deriveDevice(load('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt', '[F1 IT] V2-SAMPLE_.txt'));
const BASIC = deriveDevice(load('../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt', '[F3 FINANCE] V2-BASIC_.txt'));
const OLD = deriveDevice(load('../__fixtures__/ASHRAF-PC_.txt', 'ASHRAF-PC_.txt'));

describe('deriveDevice — full v2 scan', () => {
  it('reads the identity and the licence summary into the record', () => {
    expect(FULL.computerName).toBe('V2-SAMPLE');
    expect(FULL.scanVersion).toBe('2');
    expect(FULL.windowsLicenseType).toBe('Volume (KMS)');
    expect(FULL.officeLicenseType).toBe('Microsoft 365');
    expect(FULL.deviceJoin).toBe('Azure AD');
    expect(FULL.manufacturer).toBe('HP');
  });

  it('counts the software and keeps the raw lines for storage', () => {
    expect(FULL.softwareCount).toBe(56);
    expect(FULL.installedSoftwareRaw.split('\n')).toHaveLength(56);
    expect(FULL.installedSoftware).toHaveLength(56);
  });

  it('still derives the old fields the register already shows', () => {
    expect(FULL.installedRamGB).toBe(8);
    expect(FULL.serialNumber).toBe('5CD333KBWH');
  });

  it('exposes the parsed lists through the readers', () => {
    expect(hardwareIdsOf(FULL).memoryModules).toHaveLength(2);
    expect(officeAccountsOf(FULL)).toEqual([{ who: null, kind: 'Work or school' }]);
    expect(softwareOf(FULL)).toBe(FULL.installedSoftware);
  });
});

describe('deriveDevice — basic v2 scan', () => {
  it('says which scan wrote it', () => {
    expect(BASIC.scanVersion).toBe('2 basic');
    expect(BASIC.softwareCount).toBe(56);
    expect(BASIC.driveSerialsRaw).toBe(null);
  });
});

describe('parseReport — version 2 fixtures', () => {
  it('recognises every block, so no line is left as an unknown field or a stray value', () => {
    for (const path of ['../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt', '../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt']) {
      const parsed = parseReport(readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8'));
      expect(parsed.unknownLabels, path).toEqual([]);
      expect(parsed.warnings, path).toEqual([]);
    }
  });
});

describe('deriveDevice — old v1 report', () => {
  it('reads as version 1, with the v2 fields empty and nothing thrown', () => {
    expect(OLD.scanVersion).toBe('1');
    expect(OLD.scanRunAsAdmin).toBe(null);
    expect(OLD.manufacturer).toBe(null);
    expect(OLD.windowsLicenseType).toBe('Unknown');
    expect(OLD.officeLicenseType).toBe('Unknown');
    expect(OLD.officePersonal).toBe(null);
    expect(OLD.deviceJoin).toBe('Unknown');
    expect(OLD.softwareCount).toBe(null);
    expect(OLD.installedSoftwareRaw).toBe(null);
    expect(OLD.officeSignedInLicences).toEqual([]);
  });

  it('keeps the old fields exactly as before', () => {
    expect(OLD.computerName).toBe('ASHRAF-PC');
    expect(OLD.installedRamGB).toBe(8);
  });
});
