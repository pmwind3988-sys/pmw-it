import { describe, it, expect } from 'vitest';
import { rankIssues } from './rankIssues.js';

const group = (id, respondents, count, meanSeverity) => ({
  kind: 'bucket', id, label: id, respondents, count, meanSeverity,
});

describe('rankIssues', () => {
  it('puts five mild people above one furious person', () => {
    // The decision this whole ranking exists to encode.
    const ranked = rankIssues([
      group('lonely', 1, 5, 1),
      group('common', 5, 5, 0.1),
    ], {});
    expect(ranked[0].id).toBe('common');
  });

  it('breaks a tie on severity', () => {
    const ranked = rankIssues([
      group('calm', 4, 4, 0.1),
      group('angry', 4, 4, 0.8),
    ], {});
    expect(ranked[0].id).toBe('angry');
  });

  it('breaks a remaining tie alphabetically, so the order is stable', () => {
    const ranked = rankIssues([group('zebra', 2, 2, 0.5), group('apple', 2, 2, 0.5)], {});
    expect(ranked.map((r) => r.id)).toEqual(['apple', 'zebra']);
  });

  it('lifts pinned items to the top in pin order', () => {
    const ranked = rankIssues([
      group('big', 9, 9, 0.9),
      group('small', 1, 1, 0),
      group('mid', 4, 4, 0.4),
    ], { pinned: ['small', 'mid'] });
    expect(ranked.map((r) => r.id)).toEqual(['small', 'mid', 'big']);
    expect(ranked[0].pinned).toBe(true);
  });

  it('sinks suppressed items to the bottom without deleting them', () => {
    const ranked = rankIssues([
      group('hidden', 9, 9, 0.9),
      group('kept', 1, 1, 0),
    ], { suppressed: ['hidden'] });
    expect(ranked.map((r) => r.id)).toEqual(['kept', 'hidden']);
    expect(ranked[1].suppressed).toBe(true);
  });

  it('handles an empty list', () => {
    expect(rankIssues([], {})).toEqual([]);
  });
});
