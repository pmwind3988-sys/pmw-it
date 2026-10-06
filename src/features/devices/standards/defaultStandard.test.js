import { describe, it, expect } from 'vitest';
import {
  DEFAULT_STANDARD, defaultStandard, cloneStandard, profileKeyFor, GRADES, PARTS, PROFILE_KEYS,
} from './defaultStandard.js';
import { validateStandard } from './validateStandard.js';

describe('the default standard', () => {
  it('is valid', () => {
    expect(validateStandard(DEFAULT_STANDARD)).toEqual({ ok: true, errors: [] });
  });

  it('keeps today’s RAM floors and comfort levels', () => {
    expect(DEFAULT_STANDARD.profiles.heavy.ram).toEqual({ criticalBelow: 8, attentionBelow: 16, optimalFrom: 32 });
    expect(DEFAULT_STANDARD.profiles.desk.ram).toEqual({ criticalBelow: 8, attentionBelow: 8, optimalFrom: 16 });
    expect(DEFAULT_STANDARD.profiles.mobile.ram).toEqual({ criticalBelow: 8, attentionBelow: 16, optimalFrom: 16 });
  });

  it('asks a graphics card of engineering only', () => {
    expect(DEFAULT_STANDARD.profiles.heavy.graphics.builtIn).toBe('Needs attention');
    expect(DEFAULT_STANDARD.profiles.desk.graphics.builtIn).toBe('Optimal');
  });

  it('names exactly four grades and five parts', () => {
    expect(GRADES).toEqual(['Critical', 'Needs attention', 'Moderate', 'Optimal']);
    expect(PARTS.map((p) => p.key)).toEqual(['cpu', 'ram', 'storage', 'graphics', 'windows']);
    expect(PROFILE_KEYS).toEqual(['heavy', 'desk', 'mobile']);
  });

  it('hands out a fresh copy every time', () => {
    const a = defaultStandard();
    a.profiles.heavy.ram.optimalFrom = 99;
    expect(defaultStandard().profiles.heavy.ram.optimalFrom).toBe(32);
    const b = cloneStandard(DEFAULT_STANDARD);
    b.colors.Critical = '#000000';
    expect(DEFAULT_STANDARD.colors.Critical).toBe('#dc2626');
  });
});

describe('profileKeyFor', () => {
  it('maps a known department, whatever its case or spacing', () => {
    expect(profileKeyFor(DEFAULT_STANDARD, ' engineering ')).toBe('heavy');
    expect(profileKeyFor(DEFAULT_STANDARD, 'SALES')).toBe('mobile');
  });

  it('falls back to Desk for an unknown or missing department', () => {
    expect(profileKeyFor(DEFAULT_STANDARD, 'WAREHOUSE 9')).toBe('desk');
    expect(profileKeyFor(DEFAULT_STANDARD, null)).toBe('desk');
  });
});
