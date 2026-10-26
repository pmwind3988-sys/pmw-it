import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from './derive/deriveDevice.js';
import { identityFieldsOf, partRowsOf } from './partsView.js';

const load = (path, fileName) => ({
  text: readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8'),
  fileName,
  lastModified: Date.UTC(2026, 9, 26, 3, 28),
});

const FULL = deriveDevice(load('./__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt', '[F1 IT] V2-SAMPLE_.txt'));
const OLD = deriveDevice(load('./__fixtures__/ASHRAF-PC_.txt', 'ASHRAF-PC_.txt'));

describe('identityFieldsOf', () => {
  it('gives the six identity fields in order, with the BIOS date written out', () => {
    const fields = identityFieldsOf(FULL);
    expect(fields.map((f) => f.label)).toEqual([
      'Manufacturer', 'Product number', 'Serial number', 'Motherboard serial', 'BIOS', 'Chassis',
    ]);
    expect(fields.find((f) => f.key === 'serialNumber').value).toBe('5CD333KBWH');
    expect(fields.find((f) => f.key === 'manufacturer').value).toBe('HP');
  });

  it('leaves a value null when the scan did not read it', () => {
    const fields = identityFieldsOf(OLD);
    expect(fields.find((f) => f.key === 'productNumber').value).toBeNull();
  });
});

describe('partRowsOf', () => {
  it('lists each memory stick with its slot and part number, and the drive and screen', () => {
    const rows = partRowsOf(FULL);
    const memory = rows.filter((r) => r.part === 'Memory');
    expect(memory).toHaveLength(2);
    expect(memory[0]).toMatchObject({
      what: 'Samsung 4 GB', partNumber: 'M471A5244CB0-CWE', serial: '3430987E',
      where: 'Bottom - Slot 1 (left)',
    });
    expect(rows.some((r) => r.part === 'Drive')).toBe(true);
  });

  it('never invents a row for a battery, and is empty for an old scan', () => {
    expect(partRowsOf(FULL).some((r) => /battery/i.test(r.part))).toBe(false);
    expect(partRowsOf(OLD)).toEqual([]);
  });
});
