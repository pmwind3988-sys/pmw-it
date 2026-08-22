import { describe, it, expect } from 'vitest';
import { cosine, meanVector, assignBuckets, DEFAULT_THRESHOLD } from './similarity.js';
import { UNSORTED_ID } from './buckets.js';

const v = (...xs) => Float32Array.from(xs);

describe('cosine', () => {
  it('is 1 for the same direction and 0 for a right angle', () => {
    expect(cosine(v(1, 0), v(1, 0))).toBeCloseTo(1);
    expect(cosine(v(1, 0), v(0, 1))).toBeCloseTo(0);
  });

  it('is 0 rather than NaN for a zero vector', () => {
    expect(cosine(v(0, 0), v(1, 0))).toBe(0);
  });
});

describe('meanVector', () => {
  it('averages and re-normalises', () => {
    const mean = meanVector([v(1, 0), v(0, 1)]);
    expect(Math.hypot(mean[0], mean[1])).toBeCloseTo(1);
    expect(mean[0]).toBeCloseTo(mean[1]);
  });
});

describe('assignBuckets', () => {
  const buckets = [{ id: 'a', vector: v(1, 0) }, { id: 'b', vector: v(0, 1) }];

  it('picks the closest bucket', () => {
    const out = assignBuckets([v(0.9, 0.1), v(0.1, 0.9)], buckets, DEFAULT_THRESHOLD);
    expect(out.map((r) => r.bucketId)).toEqual(['a', 'b']);
  });

  it('declines rather than forcing a poor match', () => {
    // Equidistant and far from both: a confident wrong answer is worse
    // than an honest gap.
    const out = assignBuckets([v(0.7, 0.7)], buckets, 0.9);
    expect(out[0].bucketId).toBe(UNSORTED_ID);
  });

  it('treats the threshold as inclusive at the boundary', () => {
    const out = assignBuckets([v(1, 0)], buckets, 1);
    expect(out[0].bucketId).toBe('a');
  });

  it('returns Unsorted when there are no buckets at all', () => {
    expect(assignBuckets([v(1, 0)], [], DEFAULT_THRESHOLD)[0].bucketId).toBe(UNSORTED_ID);
  });
});
