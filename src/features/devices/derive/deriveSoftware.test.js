import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from '../parse/parseReport.js';
import { deriveSoftware, isBackgroundComponent, softwareOf } from './deriveSoftware.js';

const read = (path) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8');
const fieldsOf = (text) => parseReport(text).fields;

const FULL = fieldsOf(read('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt'));
const BASIC = fieldsOf(read('../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt'));
const OLD = fieldsOf(read('../__fixtures__/ASHRAF-PC_.txt'));

describe('deriveSoftware', () => {
  const sw = deriveSoftware(FULL);

  it('reads one entry per program, with a blank date left null', () => {
    expect(sw.installedSoftware[0]).toEqual({
      name: '4Team Safe PST Backup', version: '3.00.0682', publisher: '4Team Corporation', installedOn: '2025-10-31',
    });
    const angry = sw.installedSoftware.find((entry) => entry.name === 'Angry IP Scanner');
    expect(angry).toEqual({ name: 'Angry IP Scanner', version: '3.9.3', publisher: 'Angry IP Scanner', installedOn: null });
  });

  it('counts every entry', () => {
    expect(sw.softwareCount).toBe(sw.installedSoftware.length);
    expect(sw.softwareCount).toBe(56);
  });

  it('keeps the raw lines, one program to a line', () => {
    expect(sw.installedSoftwareRaw.split('\n')).toHaveLength(sw.softwareCount);
  });

  it('reads the basic scan with the same count', () => {
    expect(deriveSoftware(BASIC).softwareCount).toBe(sw.softwareCount);
  });

  it('gives an empty list and a null count for an old report', () => {
    expect(deriveSoftware(OLD)).toEqual({
      installedSoftware: [], installedSoftwareRaw: null, softwareCount: null,
    });
  });
});

describe('isBackgroundComponent', () => {
  it('flags the runtimes and redistributables the sample is full of', () => {
    for (const name of [
      'Microsoft Visual C++ 2008 Redistributable - x86 9.0.30729.6161',
      'Microsoft Visual C++ v14 Redistributable (x64) - 14.50.35719',
      'Microsoft .NET Runtime - 8.0.32 (x64)',
      'Microsoft Windows Desktop Runtime 10.0.13 (x86)',
      'Microsoft ASP.NET Core 8.0.32 - Shared Framework (x64)',
      'Microsoft Edge WebView2 Runtime',
    ]) {
      expect(isBackgroundComponent(name), name).toBe(true);
    }
  });

  it('flags the Windows SDK pieces', () => {
    expect(isBackgroundComponent('Windows SDK AddOn')).toBe(true);
    expect(isBackgroundComponent('Windows Software Development Kit - Windows 10.0.19041.5609')).toBe(true);
  });

  it('flags a chipset or driver package from its vendor, by publisher', () => {
    expect(isBackgroundComponent('WifiAutoInstall version 2.0.0.8', 'Realtek, Inc.')).toBe(true);
    expect(isBackgroundComponent('Intel Graphics Driver', 'Intel Corporation')).toBe(true);
  });

  it('does not flag software a person chose', () => {
    expect(isBackgroundComponent('Google Chrome', 'Google LLC')).toBe(false);
    expect(isBackgroundComponent('Microsoft Visual Studio Code (User)', 'Microsoft Corporation')).toBe(false);
    expect(isBackgroundComponent('Node.js', 'Node.js Foundation')).toBe(false);
    expect(isBackgroundComponent('Microsoft Teams Meeting Add-in for Microsoft Office', 'Microsoft')).toBe(false);
  });

  it('leaves a developer SDK and a driver removal tool alone', () => {
    expect(isBackgroundComponent('Microsoft .NET SDK 10.0.113 (x64)', 'Microsoft Corporation')).toBe(false);
    expect(isBackgroundComponent('Display Driver Uninstaller', 'Wagnardsoft')).toBe(false);
  });
});

describe('softwareOf', () => {
  it('parses a stored record, which has only the raw lines', () => {
    const stored = { installedSoftwareRaw: deriveSoftware(FULL).installedSoftwareRaw };
    expect(softwareOf(stored)).toEqual(deriveSoftware(FULL).installedSoftware);
  });

  it('returns the list of a fresh record untouched', () => {
    const fresh = deriveSoftware(FULL);
    expect(softwareOf(fresh)).toBe(fresh.installedSoftware);
  });

  it('is empty when there is nothing', () => {
    expect(softwareOf({})).toEqual([]);
  });
});
