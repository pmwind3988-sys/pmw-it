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
