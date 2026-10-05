import { describe, it, expect } from 'vitest';
import { specHistory } from './specHistory.js';

const at = (d, h = 4) => Date.UTC(2026, 2, d, h);

describe('specHistory', () => {
  const rows = specHistory([
    { fieldName: 'installedRamGB', oldValue: '8', newValue: '16', changedOn: at(3) },
    { fieldName: 'ramSlotsUsed', oldValue: '1', newValue: '2', changedOn: at(3, 5) },
    { fieldName: 'computerName', oldValue: 'FARID-HP', newValue: 'AMIR-HP', changedOn: at(1) },
    { fieldName: 'owner', oldValue: 'Farid', newValue: 'Amir', changedOn: at(1) },
    { fieldName: 'status', oldValue: 'In use', newValue: 'Spare', changedOn: at(1) },
  ]);

  it('groups by Malaysia day, newest first', () => {
    expect(rows.map((g) => g.dayLabel)).toEqual(['03/03/2026', '01/03/2026']);
    expect(rows[0].rows.map((r) => r.fieldName)).toEqual(['installedRamGB', 'ramSlotsUsed']);
  });

  it('leaves owner, department, location and status to the owner history', () => {
    expect(rows[1].rows.map((r) => r.fieldName)).toEqual(['computerName']);
  });

  it('marks a rename', () => {
    expect(rows[1].rows[0]).toMatchObject({ rename: true, oldValue: 'FARID-HP', newValue: 'AMIR-HP' });
  });

  it('labels fields the way the device page does', () => {
    expect(rows[0].rows[0].label).toBe('Installed RAM (GB)');
  });
});
