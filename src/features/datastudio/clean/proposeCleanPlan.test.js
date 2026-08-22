import { describe, it, expect } from 'vitest';
import { proposeCleanPlan } from './proposeCleanPlan.js';
import { profileDataset } from '../profile/profileDataset.js';

function planFor(grid) {
  return proposeCleanPlan(profileDataset(grid), grid);
}

describe('proposeCleanPlan', () => {
  it('proposes trimming when padded values exist, with a real count', () => {
    const grid = { headers: ['Dept'], rows: [[' HR '], ['IT '], ['HR']] };
    const step = planFor(grid).find((s) => s.op === 'trimWhitespace');
    expect(step).toMatchObject({ column: 'Dept', affectedCount: 2, enabled: true });
  });

  it('proposes a category merge and says what it will merge', () => {
    const grid = {
      headers: ['City'],
      rows: [['Kuala Lumpur'], ['kuala lumpur'], ['Penang']],
    };
    const step = planFor(grid).find((s) => s.op === 'mergeCategories');
    expect(step.preview).toContain('Kuala Lumpur');
    expect(step.affectedCount).toBe(1);
  });

  it('proposes dropping an all-null column', () => {
    const grid = { headers: ['a', 'blank'], rows: [['x', null], ['y', null]] };
    expect(planFor(grid).some((s) => s.op === 'dropEmptyColumns')).toBe(true);
  });

  it('proposes deduping only when duplicates exist', () => {
    const dupes = { headers: ['a'], rows: [['x'], ['x']] };
    const clean = { headers: ['a'], rows: [['x'], ['y']] };
    expect(planFor(dupes).some((s) => s.op === 'dedupeRows')).toBe(true);
    expect(planFor(clean).some((s) => s.op === 'dedupeRows')).toBe(false);
  });

  it('proposes nothing for an already-clean grid', () => {
    const grid = { headers: ['Dept', 'Amount'], rows: [['HR', 1], ['IT', 2]] };
    expect(planFor(grid)).toEqual([]);
  });

  it('leaves low-confidence steps disabled by default', () => {
    const grid = { headers: ['City'], rows: [['KL'], ['kl'], ['Penang']] };
    const step = planFor(grid).find((s) => s.confidence !== 'high');
    if (step) expect(step.enabled).toBe(false);
  });

  it('gives every step a unique id', () => {
    const grid = { headers: ['A', 'B'], rows: [[' x ', ' y '], ['x', 'y']] };
    const ids = planFor(grid).map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  // Spec §8.1 -- dropping rows is not obviously safe, so it is offered
  // rather than assumed. The pre-ticked/unticked split IS the safety
  // model, so it needs pinning in both directions.
  it('offers deduping unticked, since dropping rows is not obviously safe', () => {
    const grid = { headers: ['a'], rows: [['x'], ['x']] };
    const step = planFor(grid).find((s) => s.op === 'dedupeRows');
    expect(step).toMatchObject({ confidence: 'medium', enabled: false });
  });

  it('drops confidence to medium when variants differ by punctuation, not just case', () => {
    const grid = {
      headers: ['Site'],
      rows: [['PMW-SS'], ['PMW SS'], ['PMW-SS'], ['Other']],
    };
    const step = planFor(grid).find((s) => s.op === 'mergeCategories');
    expect(step).toMatchObject({ confidence: 'medium', enabled: false });
  });

  it('normalises placeholder tokens to empty and counts them', () => {
    const grid = { headers: ['Dept'], rows: [['HR'], ['N/A'], ['-']] };
    const step = planFor(grid).find((s) => s.op === 'normalizeNulls');
    expect(step).toMatchObject({ column: 'Dept', affectedCount: 2, enabled: true });
  });

  // Ruling F6 -- a cast that coerces nothing is a no-op checklist row.
  it('proposes a cast only when some value actually needs coercing', () => {
    const stored = { headers: ['Cost'], rows: [['1,234'], ['2,000'], ['3,000']] };
    const already = { headers: ['Cost'], rows: [[1234], [2000], [3000]] };
    expect(planFor(stored).some((s) => s.op === 'castType')).toBe(true);
    expect(planFor(already).some((s) => s.op === 'castType')).toBe(false);
  });

  // Later steps have to see earlier steps' output, or a merge runs
  // against values that trimming was about to change.
  it('orders steps so cleaning runs before merging before whole-grid ops', () => {
    const grid = {
      headers: ['City', 'blank'],
      rows: [[' Kuala Lumpur ', null], ['kuala lumpur', null], [' Kuala Lumpur ', null]],
    };
    const ops = planFor(grid).map((s) => s.op);
    expect(ops.indexOf('trimWhitespace')).toBeLessThan(ops.indexOf('mergeCategories'));
    expect(ops.indexOf('mergeCategories')).toBeLessThan(ops.indexOf('dropEmptyColumns'));
  });

  it('marks whole-grid steps with a null column', () => {
    const grid = { headers: ['a', 'blank'], rows: [['x', null], ['y', null]] };
    const step = planFor(grid).find((s) => s.op === 'dropEmptyColumns');
    expect(step.column).toBeNull();
  });
});
