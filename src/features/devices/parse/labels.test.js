import { describe, it, expect } from 'vitest';
import { isPlaceholder, cleanValue } from './placeholders.js';
import { KNOWN_LABELS, matchLabel } from './labels.js';

describe('isPlaceholder', () => {
  it('treats the unset SMBIOS strings as placeholders', () => {
    expect(isPlaceholder('Manufacturer1')).toBe(true);
    expect(isPlaceholder('PartNum1')).toBe(true);
    expect(isPlaceholder('System Product Name')).toBe(true);
    expect(isPlaceholder('To Be Filled By O.E.M.')).toBe(true);
    expect(isPlaceholder('Default string')).toBe(true);
  });

  it('treats None, Unknown and blanks as placeholders, case-insensitively', () => {
    expect(isPlaceholder('None')).toBe(true);
    expect(isPlaceholder('none')).toBe(true);
    expect(isPlaceholder('Unknown')).toBe(true);
    expect(isPlaceholder('   ')).toBe(true);
  });

  it('does not treat real values as placeholders', () => {
    expect(isPlaceholder('Samsung')).toBe(false);
    expect(isPlaceholder('HP Laptop 15-fd0xxx')).toBe(false);
  });
});

describe('cleanValue', () => {
  it('trims trailing whitespace the scan writes after the processor name', () => {
    expect(cleanValue('AMD Ryzen 5 7430U with Radeon Graphics         '))
      .toBe('AMD Ryzen 5 7430U with Radeon Graphics');
  });

  it('strips non-breaking and zero-width characters', () => {
    expect(cleanValue('HP\u00a0Laptop\u200b')).toBe('HP Laptop');
  });

  it('returns null for placeholders', () => {
    expect(cleanValue('None')).toBe(null);
    expect(cleanValue('')).toBe(null);
  });
});

describe('KNOWN_LABELS', () => {
  it('has the 21 labels the scan writes', () => {
    expect(KNOWN_LABELS).toHaveLength(21);
    expect(KNOWN_LABELS[0]).toBe('Name');
    expect(KNOWN_LABELS).toContain('Email data files found Active or Inactive account');
  });
});

describe('matchLabel', () => {
  it('matches a bare label with no inline value', () => {
    expect(matchLabel('Computer Name:')).toEqual({ label: 'Computer Name', inline: '' });
  });

  it('matches an inline value with a space after the colon', () => {
    expect(matchLabel('Antivirus status: NORTON ACTIVE'))
      .toEqual({ label: 'Antivirus status', inline: 'NORTON ACTIVE' });
  });

  it('matches an inline value with no space after the colon', () => {
    expect(matchLabel('Antivirus status:NORTON NOT INSTALLED'))
      .toEqual({ label: 'Antivirus status', inline: 'NORTON NOT INSTALLED' });
  });

  it('is case-insensitive and tolerates collapsed whitespace', () => {
    expect(matchLabel('total  ram:')).toEqual({ label: 'Total RAM', inline: '' });
  });

  it('does NOT match the RAM slot summary line', () => {
    expect(matchLabel('Total Slots: 2 | Used Slots: 2')).toBe(null);
  });

  it('does NOT match a mapped drive line', () => {
    expect(matchLabel('Y: | \\\\server\\PMW\\IT')).toBe(null);
  });

  it('does NOT match a value line that happens to contain a colon', () => {
    expect(matchLabel('Wi-Fi | SSID: PMW_Group | IP: 192.168.1.170 | Dynamic')).toBe(null);
  });
});
