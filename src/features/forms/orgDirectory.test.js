import { describe, it, expect } from 'vitest';
import {
  companyRows, departmentRows, companyOptions, departmentOptions, departmentStillOffered,
} from './orgDirectory.js';

const companies = companyRows([
  { Title: 'PMW Industries', Code: 'PMW', IsActive: true },
  { Title: 'PMW Lighting', Code: 'PML' },
  { Title: 'Old Co', Code: 'OLD', IsActive: false },
  { Title: 'No code', Code: '' },
]);

const departments = departmentRows([
  { Title: 'Finance', Code: 'FIN', Company: '' },
  { Title: 'Finance (Lighting)', Code: 'FIN', Company: 'PML' },
  { Title: 'Engineering', Code: 'ENG', Company: 'PMW' },
  { Title: 'Retired', Code: 'OLD', Company: '', IsActive: false },
]);

describe('company options', () => {
  it('offers active companies by name, storing the code', () => {
    expect(companyOptions(companies)).toEqual([
      { value: 'PMW', label: 'PMW Industries' },
      { value: 'PML', label: 'PMW Lighting' },
    ].sort((a, b) => a.label.localeCompare(b.label)));
  });

  it('reads a blank IsActive as active', () => {
    expect(companyOptions(companies).map((o) => o.value)).toContain('PML');
  });
});

describe('department options', () => {
  it('offers nothing until a company is picked', () => {
    expect(departmentOptions(departments, '')).toEqual([]);
  });

  it('offers shared departments plus that company\'s own', () => {
    expect(departmentOptions(departments, 'PMW').map((o) => o.value)).toEqual(['ENG', 'FIN']);
  });

  it('lets a company-specific row beat the shared one of the same code', () => {
    const fin = departmentOptions(departments, 'PML').find((o) => o.value === 'FIN');
    expect(fin.label).toBe('Finance (Lighting)');
  });

  it('never offers an inactive department, or another company\'s', () => {
    const codes = departmentOptions(departments, 'PML').map((o) => o.value);
    expect(codes).not.toContain('OLD');
    expect(codes).not.toContain('ENG');
  });

  it('matches the company case- and space-insensitively', () => {
    expect(departmentOptions(departments, ' pmw ').map((o) => o.value)).toContain('ENG');
  });
});

describe('a stale department', () => {
  it('is recognised when the company changes', () => {
    expect(departmentStillOffered(departments, 'PMW', 'ENG')).toBe(true);
    expect(departmentStillOffered(departments, 'PML', 'ENG')).toBe(false);
  });
});
