import { describe, it, expect } from 'vitest';
import {
  inStash, findInStash, searchStash, machineRow, machineTitle, machineSub,
  outOfStash, findOutOfStash, searchOutOfStash,
} from './stashMachines.js';
import { checklistFromTill } from './checklistFromTill.js';
import { IN } from '../../forms/checklistForm.js';

const devices = [
  { id: 1, computerName: 'PMW-LT-02', computerModel: 'Dell Latitude 5420', serialNumber: 'ABC123', deviceType: 'Laptop', status: 'Spare' },
  { id: 2, computerName: 'PMW-DT-01', computerModel: 'HP ProDesk 400', serialNumber: 'DT999', deviceType: 'Desktop', status: 'Spare' },
  { id: 3, computerName: 'PMW-LT-01', computerModel: 'Dell Latitude 5420', serialNumber: 'XYZ', deviceType: 'Laptop', status: 'In use', owner: 'Aina Rahman' },
  { id: 4, computerName: 'PMW-LT-03', computerModel: 'Lenovo T14', serialNumber: 'L14', deviceType: 'Laptop', status: 'Retired' },
  { id: 5, computerName: 'OLD', deviceType: 'Laptop' },
];

describe('the IT Stash at the till', () => {
  it('holds only spare machines, laptops first', () => {
    expect(inStash(devices).map((d) => d.id)).toEqual([1, 2]);
  });

  it('finds a stash machine by serial or computer name, ignoring case and spaces', () => {
    expect(findInStash(devices, ' abc 123 ')?.id).toBe(1);
    expect(findInStash(devices, 'pmw-dt-01')?.id).toBe(2);
  });

  it('does not find a machine somebody is using, or one already on the receipt', () => {
    expect(findInStash(devices, 'XYZ')).toBeNull();
    expect(findInStash(devices, 'ABC123', [devices[0]])).toBeNull();
  });

  it('searches by model, name or serial', () => {
    expect(searchStash(devices, 'latitude').map((r) => r.device.id)).toEqual([1]);
    expect(searchStash(devices, 'dt9')[0]).toMatchObject({ kind: 'machine', name: 'HP ProDesk 400 (PMW-DT-01)' });
    expect(searchStash(devices, 'l')).toEqual([]);
  });

  it('reads as a ticked laptop with its serial on the checklist', () => {
    const values = checklistFromTill({ formMode: IN, person: { name: 'Aina' }, rows: [machineRow(devices[0])] });
    expect(values.checkedItems).toEqual(['Laptop']);
    expect(values.serialNumbers).toBe('Laptop · Dell Latitude 5420 (PMW-LT-02) · S/N ABC123');
  });

  it('puts a machine of unknown type in the remarks rather than dropping it', () => {
    const row = machineRow({ computerName: 'X1', deviceType: 'Unknown', serialNumber: 'S1' });
    expect(row.category).toBe('Computer');
    const values = checklistFromTill({ formMode: IN, rows: [row] });
    expect(values.otherRemarks).toContain('X1 · S/N S1');
    expect(machineTitle({})).toBe('Machine');
  });
});

describe('taking a machine back to the stash', () => {
  it('offers machines somebody has, never spare or retired ones', () => {
    // A blank status reads as In use.
    expect(outOfStash(devices).map((d) => d.id)).toEqual([5, 3]);
  });

  it('finds one by serial or name, once', () => {
    expect(findOutOfStash(devices, 'xyz')?.id).toBe(3);
    expect(findOutOfStash(devices, 'ABC123')).toBeNull();
    expect(findOutOfStash(devices, 'XYZ', [devices[2]])).toBeNull();
  });

  it('finds one by who has it, and says so', () => {
    const [hit] = searchOutOfStash(devices, 'aina');
    expect(hit.device.id).toBe(3);
    expect(hit.sub).toBe('S/N XYZ · With Aina Rahman');
    expect(machineSub(devices[0])).toBe('S/N ABC123 · IT Stash');
  });
});
