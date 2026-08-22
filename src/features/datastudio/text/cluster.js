// The themes that are in the text whether or not anyone listed them --
// spec §6.5.
//
// Agglomerative with average linkage: start with every fragment its own
// cluster and repeatedly merge the two closest, stopping when the
// closest pair is further apart than the granularity setting. Average
// linkage rather than single linkage because single linkage chains --
// A near B, B near C, C near D -- and produces one sprawling cluster
// that means nothing.
//
// O(n^2) is fine at the scale this runs (a survey, not a corpus) and
// the guard below makes the failure loud rather than a frozen tab.

import { cosine } from './similarity.js';

export const DEFAULT_GRANULARITY = 0.45;
export const MAX_CLUSTERABLE = 5000;

export function clusterVectors(vectors, granularity = DEFAULT_GRANULARITY) {
  const list = vectors ?? [];
  if (list.length === 0) return { clusters: [], oneOffs: [] };
  if (list.length > MAX_CLUSTERABLE) {
    throw new RangeError(
      `Too many responses to group (${list.length}); the limit is ${MAX_CLUSTERABLE}.`,
    );
  }

  // Distance, not similarity: 0 is identical, 1 is unrelated.
  const distance = (a, b) => 1 - cosine(a, b);

  let groups = list.map((_, i) => [i]);

  const linkage = (left, right) => {
    let total = 0;
    for (const i of left) {
      for (const j of right) total += distance(list[i], list[j]);
    }
    return total / (left.length * right.length);
  };

  for (;;) {
    let bestDistance = Infinity;
    let bestA = -1;
    let bestB = -1;

    for (let a = 0; a < groups.length; a++) {
      for (let b = a + 1; b < groups.length; b++) {
        const d = linkage(groups[a], groups[b]);
        if (d < bestDistance) {
          bestDistance = d;
          bestA = a;
          bestB = b;
        }
      }
    }

    if (bestA === -1 || bestDistance >= granularity) break;

    const merged = [...groups[bestA], ...groups[bestB]];
    groups = groups.filter((_, i) => i !== bestA && i !== bestB);
    groups.push(merged);
  }

  const clusters = groups
    .filter((g) => g.length >= 2)
    .map((g) => g.slice().sort((a, b) => a - b))
    .sort((a, b) => b.length - a.length || a[0] - b[0]);

  // A theme of one is a quote, not a pattern. One-offs stay countable
  // and visible; they are just not presented as a finding.
  const oneOffs = groups
    .filter((g) => g.length < 2)
    .flat()
    .sort((a, b) => a - b);

  return { clusters, oneOffs };
}
