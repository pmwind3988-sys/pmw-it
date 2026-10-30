import { describe, it, expect } from 'vitest';
import { findMachines } from './finder.js';

const fleet = [
  { id: 1, computerName: 'PMW-F1-ENG-07', owner: 'Aina Rahman', serialNumber: 'SN123', anydeskId: '123 456 789' },
  { id: 2, computerName: 'PMW-F1-ENG-31', owner: 'Ali Hassan', serialNumber: 'SN999', anydeskId: '' },
  { id: 3, computerName: 'LAB-PC', owner: 'Rahman Lee', serialNumber: 'PFX7', anydeskId: '987654321' },
  { id: 4, computerName: 'Desk-9', owner: '', serialNumber: '', anydeskId: '' },
];

describe('findMachines', () => {
  it('returns nothing for a query under two characters', () => {
    expect(findMachines(fleet, '')).toEqual([]);
    expect(findMachines(fleet, ' a ')).toEqual([]);
  });

  it('finds a machine by a name it had before a rename', () => {
    const renamed = [...fleet, { id: 5, computerName: 'PMW-LT-042', previousNames: ['AMIR-HP'] }];
    expect(findMachines(renamed, 'amir-hp').map((d) => d.id)).toEqual([5]);
  });

  it('puts an exact name or serial match first', () => {
    expect(findMachines(fleet, 'lab-pc')[0].id).toBe(3);
    expect(findMachines(fleet, 'sn123')[0].id).toBe(1);
  });

  it('ranks a prefix above a mid-word match', () => {
    const ids = findMachines(fleet, 'pmw-f1').map((d) => d.id);
    expect(ids).toEqual([1, 2]);
    const rahman = findMachines(fleet, 'rahman').map((d) => d.id);
    expect(rahman).toEqual([3, 1]);
  });

  it('matches an AnyDesk ID typed without its spaces', () => {
    expect(findMachines(fleet, '123456789').map((d) => d.id)).toEqual([1]);
    expect(findMachines(fleet, '987 654 321').map((d) => d.id)).toEqual([3]);
  });

  it('is case-insensitive and respects the limit', () => {
    const many = Array.from({ length: 12 }, (_, i) => ({ id: i, computerName: `X-${i}`, owner: '' }));
    expect(findMachines(many, 'x-').length).toBe(8);
    expect(findMachines(many, 'X-', 3).length).toBe(3);
  });

  it('returns [] when nothing matches and tolerates missing fields', () => {
    expect(findMachines(fleet, 'zzzz')).toEqual([]);
    expect(findMachines([{ id: 5 }], 'ab')).toEqual([]);
    expect(findMachines(null, 'ab')).toEqual([]);
  });
});

import { findMatches, machinesWithProgram } from './finder.js';

const rich = [
  {
    id: 10, computerName: 'PMW-ENG-01', owner: 'Ali', serialNumber: 'SN-1', productNumber: '7Z8D7PA#UUF',
    memoryModules: [{ slot: '2', partNumber: 'M471A5244CB0-CWE', serial: '34309813', sizeGB: 8 }],
    driveSerials: [{ model: 'SSD', serial: 'DRV-778' }],
    installedSoftware: [{ name: 'AutoCAD 2024', version: '24', publisher: 'Autodesk', installedOn: null }],
  },
  {
    id: 11, computerName: 'AUTOCAD-PC', owner: 'Bea', serialNumber: 'SN-2',
    installedSoftware: [{ name: 'Notepad', version: '1', publisher: 'MS', installedOn: null }],
  },
  {
    id: 12, computerName: 'DESK-3', owner: 'Chen', serialNumber: 'SN-3',
    installedSoftware: [{ name: 'Google Chrome', version: '120', publisher: 'Google', installedOn: null }],
  },
  {
    id: 13, computerName: 'DESK-4', owner: 'Dee', serialNumber: 'SN-4',
    installedSoftware: [{ name: 'Google Chrome', version: '119', publisher: 'Google', installedOn: null }],
  },
  {
    id: 14, computerName: 'DESK-5', owner: 'Eli', serialNumber: 'SN-5',
    installedSoftware: [{ name: 'Google Chrome Beta', version: '1', publisher: 'Google', installedOn: null }],
  },
];

describe('findMatches reasons', () => {
  it('says why a product number matched, with the matched part', () => {
    const [match] = findMatches(rich, '7Z8D7PA');
    expect(match.device.id).toBe(10);
    expect(match.reason).toMatchObject({ label: 'Product number', value: '7Z8D7PA#UUF', at: 0, length: 7 });
  });

  it('says which memory serial and slot matched', () => {
    const [match] = findMatches(rich, '34309813');
    expect(match.device.id).toBe(10);
    expect(match.reason.label).toBe('Memory serial');
    expect(`${match.reason.label} ${match.reason.value}${match.reason.suffix}`).toBe('Memory serial 34309813 · slot 2');
  });

  it('finds a drive serial and a memory part number', () => {
    expect(findMatches(rich, 'DRV-778')[0].reason.label).toBe('Drive serial');
    expect(findMatches(rich, 'M471A5244')[0].reason.label).toBe('Memory part');
  });

  it('says a program matched, as "Has <program>"', () => {
    const [match] = findMatches(rich, 'cad 2024');
    expect(match.device.id).toBe(10);
    expect(`${match.reason.label} ${match.reason.value}`).toBe('Has AutoCAD 2024');
  });
});

describe('findMatches order', () => {
  it('a computer name match beats a program match, even when the program match is closer', () => {
    // "autocad" is the whole name of machine 11 and a program on machine 10;
    // the name wins, and the program-only machine comes after it.
    const ids = findMatches(rich, 'autocad').map((m) => m.device.id);
    expect(ids).toEqual([11, 10]);
  });

  it('keeps the tiers in order: name, owner, serial, AnyDesk, part, program', () => {
    const fleetOrder = [
      { id: 1, computerName: 'X', owner: 'Zz', serialNumber: 'QQ' },
      { id: 2, computerName: 'Y', owner: 'QQ', serialNumber: 'Zz' },
      { id: 3, computerName: 'QQ-PC', owner: 'Zz', serialNumber: 'Zz' },
    ];
    expect(findMatches(fleetOrder, 'qq').map((m) => m.device.id)).toEqual([3, 2, 1]);
  });
});

describe('machinesWithProgram', () => {
  it('counts the machines with a program whose name contains the text', () => {
    expect(machinesWithProgram(rich, 'chrome')).toBe(3);
    expect(machinesWithProgram(rich, 'excel')).toBe(0);
  });
});
