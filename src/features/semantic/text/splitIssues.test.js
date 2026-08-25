import { describe, it, expect } from 'vitest';
import { splitIssues } from './splitIssues.js';

describe('splitIssues', () => {
  it('splits a paragraph into its separate complaints', () => {
    const text = 'Financial data is collected from multiple Excel files. '
      + 'The process involves extensive manual consolidation. '
      + 'Automating extraction would reduce turnaround time.';
    expect(splitIssues(text)).toHaveLength(3);
  });

  it('strips the labels respondents paste in', () => {
    const text = 'Selected Challenge]: Data Collection\n'
      + '[Detailed Description]:\n'
      + 'I collect information from multiple WhatsApp groups and Excel files.';
    const parts = splitIssues(text);
    expect(parts.some((p) => p.includes('Detailed Description'))).toBe(false);
    expect(parts.some((p) => p.includes('WhatsApp'))).toBe(true);
  });

  it('returns nothing for a non-answer', () => {
    expect(splitIssues('no issue from IT ')).toEqual([]);
    expect(splitIssues('')).toEqual([]);
    expect(splitIssues(null)).toEqual([]);
  });

  it('does not split on an abbreviation', () => {
    const text = 'We reconcile by hand, e.g. matching invoices to receipts, every month.';
    expect(splitIssues(text)).toHaveLength(1);
  });

  it('splits on bullets and newlines', () => {
    const text = '- Approvals are chased by email\n- Nobody knows the current status\n- Reports are rebuilt from scratch';
    expect(splitIssues(text)).toHaveLength(3);
  });

  it('caps a very long answer', () => {
    const text = Array.from({ length: 30 }, (_, i) => `Problem number ${i} wastes a lot of time here.`).join(' ');
    expect(splitIssues(text).length).toBeLessThanOrEqual(12);
  });
});
