import { describe, it, expect } from 'vitest';
import { profileDataset } from './profileDataset.js';
import { profileColumn } from './profileColumn.js';

const grid = {
  headers: ['Department', 'Amount', 'Created', 'Notes'],
  rows: [
    ['HR', '100', '13/01/2024', 'first'],
    ['IT', '200', '14/01/2024', 'second'],
    ['HR', '300', '15/01/2024', ''],
  ],
};

describe('profileDataset', () => {
  it('profiles every column and counts the rows', () => {
    const p = profileDataset(grid);
    expect(p.rowCount).toBe(3);
    expect(p.columns.map((c) => c.name))
      .toEqual(['Department', 'Amount', 'Created', 'Notes']);
  });

  it('assigns the roles the canvas depends on', () => {
    const roles = Object.fromEntries(
      profileDataset(grid).columns.map((c) => [c.name, c.role]));
    expect(roles).toMatchObject({
      Department: 'dimension', Amount: 'measure', Created: 'temporal',
    });
  });

  it('computes numeric stats for measures only', () => {
    const [dept, amount] = profileDataset(grid).columns;
    expect(amount).toMatchObject({ min: 100, max: 300, mean: 200 });
    expect(dept.min).toBeNull();
  });

  it('ranks the top values for dimensions', () => {
    const dept = profileDataset(grid).columns[0];
    expect(dept.topValues[0]).toMatchObject({ value: 'HR', count: 2 });
  });

  it('picks topMeasure and primaryTemporal by non-null ratio', () => {
    const p = profileDataset(grid);
    expect(p.topMeasure).toBe('Amount');
    expect(p.primaryTemporal).toBe('Created');
  });

  it('breaks non-null ratio ties by column order', () => {
    const tied = {
      headers: ['B', 'A'],
      rows: [['1', '2'], ['3', '4']],
    };
    expect(profileDataset(tied).topMeasure).toBe('B');
  });

  it('returns null for topMeasure when there are no measures', () => {
    const noMeasures = { headers: ['Dept'], rows: [['HR'], ['IT']] };
    expect(profileDataset(noMeasures).topMeasure).toBeNull();
  });
});

describe('profileColumn overrides', () => {
  const values = ['10', '20', '30', '40'];

  it('lets the user force a type and derives the role from it', () => {
    const c = profileColumn(values, 'Amount', 0, { type: 'categorical' });
    expect(c).toMatchObject({ type: 'categorical', role: 'dimension', overridden: true });
  });

  it('recomputes the stats to match the forced type', () => {
    const asText = profileColumn(values, 'Amount', 0, { type: 'categorical' });
    // No longer a measure, so the numeric stats must go -- leaving them
    // would let a chart plot a mean for a column the user just said is
    // not a number.
    expect(asText.min).toBeNull();
    expect(asText.topValues.length).toBe(4);
  });

  it('lets the user force a role independently of the type', () => {
    const c = profileColumn(values, 'Amount', 0, { role: 'ignored' });
    expect(c).toMatchObject({ type: 'numeric', role: 'ignored' });
  });

  it('leaves the verdict alone when there is no override', () => {
    expect(profileColumn(values, 'Amount', 0).overridden).toBeUndefined();
  });
});
