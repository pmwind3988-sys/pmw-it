import { describe, it, expect } from 'vitest';
import { worstGradeOf, needsWork, gradeRank, securityRiskOf } from './health.js';

const machine = (grades) => ({
  parts: Object.fromEntries(Object.entries(grades).map(([key, grade]) => [key, { grade }])),
});

describe('worstGradeOf', () => {
  it('takes the worst part', () => {
    expect(worstGradeOf(machine({ cpu: 'Optimal', ram: 'Needs attention', storage: 'Moderate' }))).toBe('Needs attention');
    expect(worstGradeOf(machine({ cpu: 'Critical', ram: 'Optimal' }))).toBe('Critical');
  });

  it('ignores Unknown parts when another part was graded', () => {
    expect(worstGradeOf(machine({ cpu: 'Unknown', ram: 'Optimal' }))).toBe('Optimal');
  });

  it('is Unknown for a machine nobody could grade', () => {
    expect(worstGradeOf({})).toBe('Unknown');
    expect(worstGradeOf(machine({ cpu: 'Unknown' }))).toBe('Unknown');
  });
});

describe('needsWork', () => {
  it('means Critical or Needs attention on any part', () => {
    expect(needsWork(machine({ cpu: 'Critical' }))).toBe(true);
    expect(needsWork(machine({ ram: 'Needs attention' }))).toBe(true);
    expect(needsWork(machine({ ram: 'Moderate', cpu: 'Optimal' }))).toBe(false);
    expect(needsWork({})).toBe(false);
  });
});

describe('gradeRank', () => {
  it('orders worst first, Unknown last', () => {
    const sorted = ['Optimal', 'Unknown', 'Critical', 'Moderate', 'Needs attention']
      .sort((a, b) => gradeRank(a) - gradeRank(b));
    expect(sorted).toEqual(['Critical', 'Needs attention', 'Moderate', 'Optimal', 'Unknown']);
  });
});

describe('securityRiskOf', () => {
  it('never says Critical', () => {
    expect(securityRiskOf('Critical')).toBe('High');
    expect(securityRiskOf('High')).toBe('High');
    expect(securityRiskOf('Watch')).toBe('Medium');
    expect(securityRiskOf('OK')).toBe('Low');
    expect(securityRiskOf(undefined)).toBe('Unknown');
  });
});
