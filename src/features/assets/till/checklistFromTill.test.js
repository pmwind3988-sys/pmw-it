import { describe, it, expect } from 'vitest';
import { IN, OUT, INDIVIDUAL } from '../../forms/checklistForm.js';
import { cleanPreset } from '../../forms/links/linkRules.js';
import { checklistFromTill, checklistItemFor } from './checklistFromTill.js';

const ROWS = [
  { category: 'Laptop', name: 'Dell Latitude 5450', serial: '5CG4211XR2', qty: 1 },
  { category: 'Mouse', name: 'Logitech M90', serial: '', qty: 2 },
  { category: 'Phone', name: 'Samsung Galaxy A52', serial: 'R58R12ABCDE', qty: 1 },
  { category: 'Cable', name: 'HDMI cable 1.8 m', serial: '', qty: 1 },
  { category: 'Docking Station', name: 'Dell WD19S', serial: 'WD19S0042', qty: 1 },
];
const PERSON = { name: 'Aisyah Rahman', title: 'Finance Executive' };

describe('checklistItemFor', () => {
  it('puts each category where the checklist has a place for it', () => {
    expect(checklistItemFor({ category: 'Phone' }, IN)).toBe('Phone & Simcard');
    expect(checklistItemFor({ category: 'Desktop' }, IN)).toBeNull();
    expect(checklistItemFor({ category: 'Desktop' }, INDIVIDUAL)).toBe('Desktop');
    expect(checklistItemFor({ category: 'Cable', name: 'HDMI cable' }, INDIVIDUAL)).toBe('HDMI Cable');
    expect(checklistItemFor({ category: 'Cable', name: 'HDMI cable' }, IN)).toBeNull();
    expect(checklistItemFor({ category: 'Docking Station' }, IN)).toBeNull();
  });
});

describe('checklistFromTill', () => {
  it('ticks the boxes, writes every serial, and keeps the rest in remarks (IN)', () => {
    const values = checklistFromTill({ formMode: IN, person: PERSON, rows: ROWS, date: '2026-10-07' });
    expect(values).toMatchObject({
      formMode: IN, employeeName: 'Aisyah Rahman', position: 'Finance Executive', formDate: '2026-10-07',
      employeeNo: '', entity: '', department: '',
      checkedItems: ['Laptop', 'Mouse', 'Phone & Simcard'],
    });
    expect(values.serialNumbers.split('\n')).toEqual([
      'Laptop · Dell Latitude 5450 · S/N 5CG4211XR2',
      'Phone · Samsung Galaxy A52 · S/N R58R12ABCDE',
      'Docking Station · Dell WD19S · S/N WD19S0042',
    ]);
    expect(values.otherRemarks).toBe('Also handed over:\nHDMI cable 1.8 m\nDell WD19S · S/N WD19S0042');
  });

  it('lists item × quantity for an individual request', () => {
    const values = checklistFromTill({ formMode: INDIVIDUAL, person: PERSON, rows: ROWS });
    expect(values.items).toEqual([
      { item: 'Laptop', quantity: 1 },
      { item: 'Mouse', quantity: 2 },
      { item: 'Phone & Simcard', quantity: 1 },
      { item: 'HDMI Cable', quantity: 1 },
    ]);
    expect(values.otherRemarks).toBe('Also handed over:\nDell WD19S · S/N WD19S0042');
  });

  it('says "returned" on an OUT checklist', () => {
    const values = checklistFromTill({ formMode: OUT, person: PERSON, rows: [ROWS[4]] });
    expect(values.otherRemarks.startsWith('Also returned:')).toBe(true);
    expect(values.checkedItems).toEqual([]);
  });

  it('survives the link’s own clean-up unchanged', () => {
    const values = checklistFromTill({ formMode: IN, person: PERSON, rows: ROWS });
    const preset = cleanPreset(values, null);
    expect(preset.checkedItems).toEqual(values.checkedItems);
    expect(preset.serialNumbers).toBe(values.serialNumbers);
    expect(preset.otherRemarks).toBe(values.otherRemarks);
  });
});
