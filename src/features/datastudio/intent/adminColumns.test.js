import { describe, it, expect } from 'vitest';
import { detectAdminColumns, normalizeHeader } from './adminColumns.js';

// A profile stand-in. Only `name`, `index` and `role` are read, so the
// rest of a real column's shape would be noise here.
function profileOf(names, role = 'dimension') {
  return { columns: names.map((name, index) => ({ name, index, role })) };
}

describe('normalizeHeader', () => {
  it('flattens punctuation and case', () => {
    expect(normalizeHeader('Submitted_at:')).toBe('submitted at');
    expect(normalizeHeader('  Time Taken (s) ')).toBe('time taken s');
  });
});

describe('detectAdminColumns', () => {
  const hidden = (names, role) => detectAdminColumns(profileOf(names, role)).map((c) => c.name);

  it('catches the timing columns a form adds', () => {
    expect(hidden(['Timestamp', 'Completion time', 'Date answered', 'Last modified']))
      .toEqual(['Timestamp', 'Completion time', 'Date answered', 'Last modified']);
  });

  it('catches how long the form took', () => {
    expect(hidden(['Time taken', 'Duration (minutes)'], 'measure'))
      .toEqual(['Time taken', 'Duration (minutes)']);
  });

  it('catches who answered', () => {
    expect(hidden(['Email Address', 'Respondent ID', 'IP address']))
      .toEqual(['Email Address', 'Respondent ID', 'IP address']);
  });

  it('catches running numbers and unlabelled columns', () => {
    expect(hidden(['No.', 'Bil', 'Unnamed: 3', ''], 'measure'))
      .toEqual(['No.', 'Bil', 'Unnamed: 3', '']);
  });

  it('leaves the questions the survey actually asked', () => {
    expect(hidden([
      'Department', 'Describe the biggest issue', 'How often does it happen',
      'Satisfaction score', 'Date of incident', 'Hours lost per week',
    ])).toEqual([]);
  });

  it('does not hide a date just for being a date', () => {
    // `Date of incident` is a real answer and `Timestamp` is not, and
    // nothing about their values separates them.
    expect(hidden(['Date of incident', 'Timestamp'], 'temporal')).toEqual(['Timestamp']);
  });

  it('skips columns the profiler already parked', () => {
    expect(hidden(['Timestamp', 'Email'], 'ignored')).toEqual([]);
  });

  it('reports why each column was picked', () => {
    const [first] = detectAdminColumns(profileOf(['Time taken'], 'measure'));
    expect(first.reason).toMatch(/how long/);
    expect(first.index).toBe(0);
  });

  it('survives a missing profile', () => {
    expect(detectAdminColumns(null)).toEqual([]);
  });
});
