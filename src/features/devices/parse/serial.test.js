import { describe, it, expect } from 'vitest';
import { cleanSerial, isPlaceholderSerial, normaliseSerial } from './placeholders.js';

describe('normaliseSerial', () => {
  it('upper-cases and removes every space', () => {
    expect(normaliseSerial(' 5cg8241 kqz ')).toBe('5CG8241KQZ');
  });

  it('removes invisible characters too', () => {
    expect(normaliseSerial('5CG\u00a08241\u200bKQZ')).toBe('5CG8241KQZ');
  });
});

describe('cleanSerial', () => {
  it('keeps a real serial', () => {
    expect(cleanSerial('MXL7180Q2B')).toBe('MXL7180Q2B');
  });

  it.each([
    'System Serial Number', 'Chassis Serial Number', 'Default string',
    'To be filled by O.E.M.', 'Not specified', '0', '00000000', 'XXXXXXXX', '', '   ', '\u00a0', null, undefined,
  ])('treats %j as no serial', (value) => {
    expect(cleanSerial(value)).toBeNull();
    expect(isPlaceholderSerial(value)).toBe(true);
  });

  it('does not treat a serial that merely contains a repeat as junk', () => {
    expect(cleanSerial('5CG000001')).toBe('5CG000001');
  });
});
