import { describe, it, expect } from 'vitest';
import { stripLabelPrefix, isNonAnswer, normalizeText } from './boilerplate.js';

describe('stripLabelPrefix', () => {
  it('removes the labels respondents copy out of the question', () => {
    expect(stripLabelPrefix('[Selected Challenge]: Data Collection')).toBe('Data Collection');
    expect(stripLabelPrefix('[Detailed Description]:')).toBe('');
    expect(stripLabelPrefix('Description: we retype everything')).toBe('we retype everything');
    // The real export is missing the opening bracket on some rows.
    expect(stripLabelPrefix('Selected Challenge]: Data Collection')).toBe('Data Collection');
  });

  it('leaves an ordinary sentence alone', () => {
    expect(stripLabelPrefix('The problem: nobody owns the report'))
      .toBe('The problem: nobody owns the report');
  });
});

describe('isNonAnswer', () => {
  it('drops the ways people say nothing is wrong', () => {
    expect(isNonAnswer('no issue from IT')).toBe(true);
    expect(isNonAnswer('N/A')).toBe(true);
    expect(isNonAnswer('-')).toBe(true);
    expect(isNonAnswer('none')).toBe(true);
    expect(isNonAnswer('')).toBe(true);
  });

  it('keeps a real complaint that starts with "no"', () => {
    // This pair is the whole rule. A prefix match alone deletes it.
    expect(isNonAnswer('No proper system exists for tracking approvals')).toBe(false);
    expect(isNonAnswer('Nothing is documented, so every handover starts over')).toBe(false);
  });
});

describe('normalizeText', () => {
  it('strips zero-width characters and collapses whitespace', () => {
    expect(normalizeText('a\u200Bb   c\uFEFF')).toBe('ab c');
  });

  it('survives a non-string', () => {
    expect(normalizeText(null)).toBe('');
    expect(normalizeText(42)).toBe('42');
  });
});
