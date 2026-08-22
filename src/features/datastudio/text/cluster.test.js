import { describe, it, expect } from 'vitest';
import { clusterVectors, DEFAULT_GRANULARITY, MAX_CLUSTERABLE } from './cluster.js';

const v = (...xs) => Float32Array.from(xs);

describe('clusterVectors', () => {
  it('groups vectors that point the same way', () => {
    const vectors = [
      v(1, 0), v(0.99, 0.1), v(0.98, 0.05),
      v(0, 1), v(0.1, 0.99),
    ];
    const { clusters } = clusterVectors(vectors, DEFAULT_GRANULARITY);
    expect(clusters).toHaveLength(2);
    expect(clusters[0]).toHaveLength(3);
    expect(clusters[0]).toEqual(expect.arrayContaining([0, 1, 2]));
  });

  it('keeps a lone vector out of the themes', () => {
    // A theme of one is a quote, not a pattern.
    const vectors = [v(1, 0), v(0.99, 0.1), v(0, 1)];
    const { clusters, oneOffs } = clusterVectors(vectors, DEFAULT_GRANULARITY);
    expect(clusters).toHaveLength(1);
    expect(oneOffs).toEqual([2]);
  });

  it('sweeps up more of the stragglers as granularity rises', () => {
    // Cluster COUNT is not monotonic in granularity -- at a tight
    // setting everything is a one-off and there are no clusters at all,
    // so counting them would read as 'broader made more themes'. What
    // does move in one direction is how much is left ungrouped.
    const vectors = [v(1, 0), v(0.9, 0.44), v(0.44, 0.9), v(0, 1)];
    const narrow = clusterVectors(vectors, 0.1);
    const broad = clusterVectors(vectors, 0.9);
    expect(broad.oneOffs.length).toBeLessThan(narrow.oneOffs.length);
    expect(narrow.clusters).toHaveLength(0);
    expect(broad.clusters).toHaveLength(1);
  });

  it('returns nothing for an empty input', () => {
    expect(clusterVectors([], DEFAULT_GRANULARITY)).toEqual({ clusters: [], oneOffs: [] });
  });

  it('refuses rather than freezing on an enormous input', () => {
    const huge = Array.from({ length: MAX_CLUSTERABLE + 1 }, () => v(1, 0));
    expect(() => clusterVectors(huge, DEFAULT_GRANULARITY)).toThrow(RangeError);
  });
});
