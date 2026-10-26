import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from '../parse/parseReport.js';
import { deriveHardwareIds, hardwareIdsOf } from './deriveHardwareIds.js';

const read = (path) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8');
const fieldsOf = (text) => parseReport(text).fields;

const FULL = fieldsOf(read('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt'));
const BASIC = fieldsOf(read('../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt'));
const OLD = fieldsOf(read('../__fixtures__/ASHRAF-PC_.txt'));

describe('deriveHardwareIds — full v2 scan', () => {
  const hw = deriveHardwareIds(FULL);

  it('reads the maker and the numbers the maker set', () => {
    expect(hw.manufacturer).toBe('HP');
    expect(hw.productNumber).toBe('7Z8D7PA#UUF');
    expect(hw.deviceUuid).toBe('33444335-3333-424B-5748-484233334435');
    expect(hw.modelFamily).toBe(null); // the scan wrote the label with nothing after it
  });

  it('cleans the motherboard serial as a serial', () => {
    expect(hw.motherboardSerial).toBe('PRSPC058JIG0O2');
  });

  it('reads the BIOS as a vendor and version, and its date', () => {
    expect(hw.biosVersion).toBe('AMI F.10');
    expect(hw.biosDate).toBe('2024-03-20');
  });

  it('reads the chassis kind and leaves a masked asset tag null', () => {
    expect(hw.chassis).toBe('Notebook');
    expect(hw.assetTag).toBe(null);
  });

  it('reads the Windows build as one readable line', () => {
    expect(hw.windowsBuild).toBe('26H2 (build 26300.9711, 64-bit)');
  });

  it('parses each memory stick into its parts', () => {
    expect(hw.memoryModules).toEqual([
      { slot: 'Bottom - Slot 1 (left)', sizeGB: 4, maker: 'Samsung', partNumber: 'M471A5244CB0-CWE', serial: '3430987E' },
      { slot: 'Bottom - Slot 2 (right)', sizeGB: 4, maker: 'Samsung', partNumber: 'M471A5244CB0-CWE', serial: '34309813' },
    ]);
  });

  it('keeps the memory lines as raw text for storage', () => {
    expect(hw.memoryModulesRaw.split('\n')).toHaveLength(2);
    expect(hw.memoryModulesRaw.startsWith('Bottom - Slot 1 (left) | 4 GB')).toBe(true);
  });

  it('parses the drive, its bus and firmware, and keeps the serial as written', () => {
    expect(hw.driveSerials).toEqual([{
      model: 'KBG50ZNV512G KIOXIA',
      serial: '0000_0000_0000_0000_8CE3_8E04_044E_3F10',
      bus: 'NVMe',
      firmware: 'HP02AN00',
    }]);
  });

  it('reads a monitor with no serial, and treats week 0 as no week', () => {
    expect(hw.monitorDetails).toEqual([{
      maker: 'AUO', name: null, serial: null, madeYear: 2022, madeWeek: null,
    }]);
  });
});

describe('deriveHardwareIds — basic v2 scan (no wmic)', () => {
  const hw = deriveHardwareIds(BASIC);

  it('reads what the cmd-only scan did write', () => {
    expect(hw.manufacturer).toBe('HP');
    expect(hw.biosVersion).toBe('AMI F.10');
    expect(hw.windowsBuild).toBe('26H2 (build 26300.9711, 64-bit)');
  });

  it('gives nulls and empty lists for the blocks the basic scan left empty', () => {
    expect(hw.motherboardSerial).toBe(null);
    expect(hw.deviceUuid).toBe(null);
    expect(hw.chassis).toBe(null);
    expect(hw.memoryModules).toEqual([]);
    expect(hw.memoryModulesRaw).toBe(null);
    expect(hw.driveSerials).toEqual([]);
    expect(hw.monitorDetails).toEqual([]);
  });
});

describe('deriveHardwareIds — old v1 report', () => {
  it('gives nulls and empty lists, and does not throw', () => {
    const hw = deriveHardwareIds(OLD);
    expect(hw.manufacturer).toBe(null);
    expect(hw.biosVersion).toBe(null);
    expect(hw.biosDate).toBe(null);
    expect(hw.memoryModules).toEqual([]);
    expect(hw.windowsBuild).toBe(null);
  });
});

describe('hardwareIdsOf', () => {
  it('parses a stored record that only has the raw text', () => {
    const stored = deriveHardwareIds(FULL);
    const { memoryModulesRaw, driveSerialsRaw, monitorDetailsRaw } = stored;
    const back = hardwareIdsOf({ memoryModulesRaw, driveSerialsRaw, monitorDetailsRaw });
    expect(back.memoryModules).toEqual(stored.memoryModules);
    expect(back.driveSerials).toEqual(stored.driveSerials);
    expect(back.monitorDetails).toEqual(stored.monitorDetails);
  });

  it('returns the arrays of a fresh record untouched', () => {
    const fresh = deriveHardwareIds(FULL);
    expect(hardwareIdsOf(fresh).memoryModules).toBe(fresh.memoryModules);
  });

  it('is empty for a record with none of them', () => {
    expect(hardwareIdsOf({})).toEqual({ memoryModules: [], driveSerials: [], monitorDetails: [] });
  });
});
