import { describe, expect, it } from 'vitest';
import {
  scanKindOf, officeWarningOf, memoryPartsOf, driveSerialsOf, programsMatching, hasSoftware,
} from './fleetFacts.js';

describe('scanKindOf', () => {
  it('reads new, basic and old, and treats a missing version as old', () => {
    expect(scanKindOf('2')).toBe('new');
    expect(scanKindOf('2 basic')).toBe('basic');
    expect(scanKindOf('1')).toBe('old');
    expect(scanKindOf(null)).toBe('old');
    expect(scanKindOf('')).toBe('old');
  });
});

describe('officeWarningOf', () => {
  it('is blank unless the machine is flagged as personal', () => {
    expect(officeWarningOf({ officePersonal: false })).toBeNull();
    expect(officeWarningOf({ officePersonal: null })).toBeNull();
  });

  it('names a personal copy, a personal account, or both', () => {
    expect(officeWarningOf({ officePersonal: true, officePersonalReason: 'A personal Microsoft 365 (O365HomePremRetail) is installed' }))
      .toBe('Personal copy');
    expect(officeWarningOf({ officePersonal: true, officePersonalReason: 'A personal Microsoft account (a@b.com) is signed in to Office on this machine' }))
      .toBe('Personal account');
    expect(officeWarningOf({
      officePersonal: true,
      officePersonalReason: 'A personal Microsoft 365 (X) is installed; A personal Microsoft account is signed in to Office on this machine',
    })).toBe('Personal copy and account');
  });
});

describe('memoryPartsOf and driveSerialsOf', () => {
  const device = {
    memoryModules: [
      { slot: '1', maker: 'Micron', partNumber: 'M471A5244CB0-CWE', serial: '34309813' },
      { slot: '2', maker: 'Micron', partNumber: 'M471A5244CB0-CWE', serial: '34309814' },
    ],
    driveSerials: [{ model: 'SAMSUNG MZVLB512', serial: 'S4EWNX0', bus: 'NVMe' }, { model: 'WDC', serial: null }],
  };

  it('groups identical memory parts with a count', () => {
    expect(memoryPartsOf(device)).toEqual(['2 × M471A5244CB0-CWE']);
  });

  it('lists every drive serial, and says so when a drive gave none', () => {
    expect(driveSerialsOf(device)).toEqual(['S4EWNX0', 'WDC (no serial)']);
  });
});

describe('programsMatching', () => {
  const device = {
    installedSoftware: [
      { name: 'AutoCAD 2024', version: '24.0', publisher: 'Autodesk', installedOn: null },
      { name: 'Google Chrome', version: '120', publisher: 'Google', installedOn: null },
    ],
  };

  it('matches part of a program name, case-insensitively', () => {
    expect(programsMatching(device, 'autocad').map((p) => p.name)).toEqual(['AutoCAD 2024']);
    expect(hasSoftware(device, 'CHROME')).toBe(true);
    expect(hasSoftware(device, 'excel')).toBe(false);
  });

  it('matches nothing for an empty text', () => {
    expect(programsMatching(device, '  ')).toEqual([]);
  });
});
