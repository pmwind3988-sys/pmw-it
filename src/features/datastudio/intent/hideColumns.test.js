import { describe, it, expect } from 'vitest';
import { hideColumns, unhideColumns } from './hideColumns.js';

const profile = {
  rowCount: 3,
  topMeasure: 'Time taken',
  primaryTemporal: 'Timestamp',
  columns: [
    { name: 'Timestamp', index: 0, type: 'datetime', role: 'temporal', distinctCount: 3, nonNullRatio: 1 },
    { name: 'Department', index: 1, type: 'categorical', role: 'dimension', distinctCount: 2, nonNullRatio: 1 },
    { name: 'Time taken', index: 2, type: 'numeric', role: 'measure', distinctCount: 3, nonNullRatio: 1 },
    { name: 'Hours lost', index: 3, type: 'numeric', role: 'measure', distinctCount: 3, nonNullRatio: 0.9 },
  ],
};

const hidden = [
  { name: 'Timestamp', index: 0, reason: 'when the form was filled in' },
  { name: 'Time taken', index: 2, reason: 'how long the form took' },
];

describe('hideColumns', () => {
  const result = hideColumns(profile, hidden);

  it('parks the named column and leaves the rest alone', () => {
    expect(result.columns[0].role).toBe('ignored');
    expect(result.columns[1].role).toBe('dimension');
  });

  it('keeps the type and stats, so nothing has to be re-measured', () => {
    expect(result.columns[0].type).toBe('datetime');
    expect(result.columns[0].distinctCount).toBe(3);
  });

  it('marks the column as decided rather than inferred', () => {
    expect(result.columns[0].overridden).toBe(true);
  });

  it('does not mutate the profile it was given', () => {
    expect(profile.columns[0].role).toBe('temporal');
  });

  // The visible symptom of getting this wrong: the starter charts open
  // with "Time taken over Timestamp" on a sheet where both were parked.
  it('stops naming a hidden column as the headline measure', () => {
    expect(result.topMeasure).toBe('Hours lost');
    expect(result.primaryTemporal).toBeNull();
  });

  it('returns the profile untouched when there is nothing to hide', () => {
    expect(hideColumns(profile, [])).toBe(profile);
  });
});

describe('unhideColumns', () => {
  it('restores the role the type implies', () => {
    const back = unhideColumns(hideColumns(profile, hidden), hidden);
    expect(back.columns[0].role).toBe('temporal');
    expect(back.columns[0].overridden).toBe(false);
  });

  it('gives the headline measure and temporal back too', () => {
    const back = unhideColumns(hideColumns(profile, hidden), hidden);
    expect(back.topMeasure).toBe('Time taken');
    expect(back.primaryTemporal).toBe('Timestamp');
  });

  it('round-trips every type the profiler produces', () => {
    const all = {
      columns: [
        { name: 'a', type: 'numeric', role: 'measure' },
        { name: 'b', type: 'categorical', role: 'dimension' },
        { name: 'c', type: 'date', role: 'temporal' },
        { name: 'd', type: 'identifier', role: 'ignored' },
      ],
    };
    const marks = all.columns.map((c) => ({ name: c.name }));
    const roles = unhideColumns(hideColumns(all, marks), marks).columns.map((c) => c.role);
    expect(roles).toEqual(['measure', 'dimension', 'temporal', 'ignored']);
  });
});
