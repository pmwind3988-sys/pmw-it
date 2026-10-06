import { describe, it, expect } from 'vitest';
import { inkFor, tooSimilar, gradeCssVars, contrastRatio } from './gradeColors.js';
import { DEFAULT_COLORS } from './defaultStandard.js';

describe('gradeColors', () => {
  it('puts dark text on light colours and white text on dark ones', () => {
    expect(inkFor('#ffffff')).toBe('#101828');
    expect(inkFor('#101828')).toBe('#ffffff');
    expect(inkFor('#f59e0b')).toBe('#101828');
  });

  it('measures contrast like WCAG', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 0);
    expect(contrastRatio('#123456', '#123456')).toBeCloseTo(1, 5);
  });

  it('warns only when two grades are too alike by colour distance', () => {
    expect(tooSimilar(DEFAULT_COLORS)).toBe(false);
    expect(tooSimilar({ ...DEFAULT_COLORS, Moderate: '#12a150' })).toBe(true);
    expect(tooSimilar({ ...DEFAULT_COLORS, Optimal: '#14a352' })).toBe(false);
    expect(tooSimilar({ ...DEFAULT_COLORS, Optimal: '#f5a00c' })).toBe(true);
  });

  it('turns the colours into CSS variables with a text colour each', () => {
    const vars = gradeCssVars(DEFAULT_COLORS);
    expect(vars['--grade-critical']).toBe('#dc2626');
    expect(vars['--grade-attention-ink']).toBe('#101828');
    expect(Object.keys(vars)).toHaveLength(10);
  });
});
