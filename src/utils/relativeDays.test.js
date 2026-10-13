import { describe, it, expect } from 'vitest';
import { relativeDays } from './relativeDays.js';

const NOON = new Date(2026, 9, 13, 12, 0, 0).getTime();
const daysBefore = (n) => new Date(2026, 9, 13 - n, 9, 30, 0).getTime();

describe('relativeDays', () => {
  it('reads the same calendar day as today', () => {
    expect(relativeDays(new Date(2026, 9, 13, 1, 0).getTime(), NOON)).toBe('today');
  });

  it('reads one calendar day back as yesterday', () => {
    expect(relativeDays(daysBefore(1), NOON)).toBe('yesterday');
  });

  it('counts days under two months', () => {
    expect(relativeDays(daysBefore(9), NOON)).toBe('9 days ago');
    expect(relativeDays(daysBefore(59), NOON)).toBe('59 days ago');
  });

  it('counts months beyond that', () => {
    expect(relativeDays(daysBefore(120), NOON)).toBe('3 months ago');
  });

  it('is empty for anything that is not an instant', () => {
    expect(relativeDays(undefined, NOON)).toBe('');
    expect(relativeDays(NaN, NOON)).toBe('');
    expect(relativeDays('2026-10-01', NOON)).toBe('');
  });
});
