import { describe, it, expect } from 'vitest';
import { applyCleanPlan } from './applyCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';
import { proposeCleanPlan } from './proposeCleanPlan.js';

// RULING (amends the plan's fixture, not its assertions): the plan gave
// this grid three rows, one of which put 'nope' in Amount. That is 2/3
// numeric, and spec §7.2's 95% rule types it CATEGORICAL -- so
// `expect(col.values).toBeInstanceOf(Float64Array)` could never pass
// against the plan's own rules. The spec wins over the plan, so the grid
// is widened to twenty rows: nineteen numbers and the same single 'nope'
// is exactly 95%, which keeps Amount numeric AND keeps a real coercion
// failure to assert NaN on. Every assertion below is the plan's, intact.
const EXTRA_DEPTS = ['IT', 'HR', 'Finance'];

const grid = {
  headers: ['Dept', 'Amount', 'Created'],
  rows: [
    [' HR ', '1,000', '13/01/2024'],
    ['hr', '2,000', '14/01/2024'],
    ['IT', 'nope', '15/01/2024'],
    ...Array.from({ length: 17 }, (_, i) => [
      EXTRA_DEPTS[i % 3],
      `${i + 3},000`,
      `${String((i % 28) + 1).padStart(2, '0')}/02/2024`,
    ]),
  ],
};

function build(enabledOverride) {
  const profile = profileDataset(grid);
  let plan = proposeCleanPlan(profile, grid);
  if (enabledOverride) plan = plan.map(enabledOverride);
  return applyCleanPlan(grid, plan, profile);
}

describe('applyCleanPlan', () => {
  it('stores numerics as a Float64Array with NaN for failures', () => {
    const col = build().columns.find((c) => c.name === 'Amount');
    expect(col.values).toBeInstanceOf(Float64Array);
    expect(col.values[0]).toBe(1000);
    expect(col.values[2]).toBeNaN();
  });

  it('dictionary-encodes categoricals', () => {
    const col = build().columns.find((c) => c.name === 'Dept');
    expect(col.values).toBeInstanceOf(Int32Array);
    expect(col.dictionary).toContain('HR');
    // ' HR ' and 'hr' both normalise onto the same code.
    expect(col.values[0]).toBe(col.values[1]);
  });

  it('stores temporal columns as epoch ms', () => {
    const col = build().columns.find((c) => c.name === 'Created');
    expect(col.values).toBeInstanceOf(Float64Array);
    expect(col.values[0]).toBe(Date.UTC(2024, 0, 13));
  });

  it('exposes columns by name', () => {
    expect(build().byName.get('Amount')).toBeDefined();
  });

  // Spec §8.1 -- the plan is data, so disabling a step must change the output.
  it('respects disabled steps', () => {
    const withMerge = build();
    const withoutMerge = build((s) => (
      s.op === 'mergeCategories' || s.op === 'unifyCase' ? { ...s, enabled: false } : s));
    const a = withMerge.columns.find((c) => c.name === 'Dept');
    const b = withoutMerge.columns.find((c) => c.name === 'Dept');
    expect(b.dictionary.length).toBeGreaterThan(a.dictionary.length);
  });

  it('is idempotent -- applying the same plan twice gives the same result', () => {
    const a = build();
    const b = build();
    expect(Array.from(a.columns[1].values)).toEqual(Array.from(b.columns[1].values));
  });

  it('never mutates the input grid', () => {
    const snapshot = JSON.stringify(grid);
    build();
    expect(JSON.stringify(grid)).toBe(snapshot);
  });

  // --- the null encodings every later phase reads (spec §6.1) ---------
  //
  // These are a contract, not an implementation detail: filterMask and
  // aggregate both test for them by value, so a change here silently
  // turns "missing" into a real category or a real zero.

  it('encodes a missing numeric as NaN, not zero', () => {
    const col = build().columns.find((c) => c.name === 'Amount');
    expect(col.values[2]).toBeNaN();
    expect(col.values[2]).not.toBe(0);
  });

  it('encodes a missing category as -1, outside the dictionary', () => {
    // Two columns on purpose: in a one-column grid the null row is an
    // ENTIRELY empty row, which dropEmptyRows correctly removes before
    // the encoder ever sees it.
    const sparse = { headers: ['Dept', 'Site'], rows: [['HR', 'a'], [null, 'b'], ['IT', 'c']] };
    const profile = profileDataset(sparse);
    const ds = applyCleanPlan(sparse, proposeCleanPlan(profile, sparse), profile);
    expect(ds.columns[0].values[1]).toBe(-1);
  });

  it('encodes a missing boolean as 2, distinct from both true and false', () => {
    const bools = {
      headers: ['Active', 'Site'],
      rows: [['Yes', 'a'], ['No', 'b'], [null, 'c']],
    };
    const profile = profileDataset(bools);
    const ds = applyCleanPlan(bools, proposeCleanPlan(profile, bools), profile);
    expect(Array.from(ds.columns[0].values)).toEqual([1, 0, 2]);
    expect(ds.columns[0].values).toBeInstanceOf(Uint8Array);
  });

  it('keeps text columns as plain strings with null for missing', () => {
    const notes = {
      headers: ['Notes', 'Site'],
      rows: Array.from({ length: 60 }, (_, i) => [i === 0 ? null : `remark ${i}`, 'a']),
    };
    const profile = profileDataset(notes);
    const ds = applyCleanPlan(notes, proposeCleanPlan(profile, notes), profile);
    expect(ds.columns[0].values[0]).toBeNull();
    expect(ds.columns[0].values[1]).toBe('remark 1');
    expect(ds.columns[0].dictionary).toBeNull();
  });

  it('reports the row count and carries the column roles through', () => {
    const ds = build();
    expect(ds.rowCount).toBe(20);
    const roles = Object.fromEntries(ds.columns.map((c) => [c.name, c.role]));
    expect(roles).toMatchObject({
      Dept: 'dimension', Amount: 'measure', Created: 'temporal',
    });
  });

  it('drops rows and columns that whole-grid steps remove', () => {
    const messy = {
      headers: ['Dept', 'blank'],
      rows: [['HR', null], ['HR', null], [null, null]],
    };
    const profile = profileDataset(messy);
    const plan = proposeCleanPlan(profile, messy)
      // Deduping is offered unticked, so tick it explicitly here.
      .map((s) => ({ ...s, enabled: true }));
    const ds = applyCleanPlan(messy, plan, profile);
    expect(ds.columns.map((c) => c.name)).toEqual(['Dept']);
    expect(ds.rowCount).toBe(1);
  });

  it('marks a percent column so charts can format it', () => {
    const rates = {
      headers: ['Rate', 'Site'],
      rows: Array.from({ length: 6 }, (_, i) => [`${40 + i}%`, 'a']),
    };
    const profile = profileDataset(rates);
    const ds = applyCleanPlan(rates, proposeCleanPlan(profile, rates), profile);
    expect(ds.columns[0].isPercent).toBe(true);
    expect(ds.columns[0].values[0]).toBeCloseTo(0.4);
  });
});
