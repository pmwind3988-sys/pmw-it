// The themes that are in the text whether or not anyone listed them --
// spec §6.5.
//
// Agglomerative with average linkage (UPGMA): start with every fragment
// its own cluster and repeatedly merge the two closest, stopping when
// the closest pair is further apart than the granularity setting.
// Average linkage rather than single linkage because single linkage
// chains -- A near B, B near C, C near D -- and produces one sprawling
// cluster that means nothing.
//
// The distances are computed ONCE into a matrix and then updated in
// place by the Lance-Williams rule as clusters merge:
//
//   d(A + B, C) = (|A| * d(A,C) + |B| * d(B,C)) / (|A| + |B|)
//
// The obvious implementation instead re-sums every member-to-member
// distance on every merge round, which is O(n^3) overall. Measured on
// the real survey -- 134 fragments -- that took between 3 and 10 seconds
// depending on the setting, against a 300ms budget for a control the
// user drags. This version is O(n^2).

import { cosine } from './similarity.js';

// Chosen against the real survey rather than guessed. At 0.45 only 54 of
// 134 fragments joined a theme at all and 80 were left as one-offs; at
// 0.75 a single theme swallowed 62 of them and stopped meaning anything.
// 0.65 groups 123 of 134 with the largest theme at 19 -- most of the
// corpus placed, nothing dominating it.
export const DEFAULT_GRANULARITY = 0.65;
export const MAX_CLUSTERABLE = 5000;

export function clusterVectors(vectors, granularity = DEFAULT_GRANULARITY) {
  const list = vectors ?? [];
  const n = list.length;
  if (n === 0) return { clusters: [], oneOffs: [] };
  if (n > MAX_CLUSTERABLE) {
    throw new RangeError(
      `Too many responses to group (${n}); the limit is ${MAX_CLUSTERABLE}.`,
    );
  }

  // Distance, not similarity: 0 is identical, 1 is unrelated.
  const distance = new Float64Array(n * n);
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      const d = 1 - cosine(list[i], list[j]);
      distance[i * n + j] = d;
      distance[j * n + i] = d;
    }
  }

  // Each surviving cluster keeps its members and its row in the matrix;
  // a merged-away cluster is marked dead rather than spliced out, so no
  // index ever has to be rewritten.
  const members = Array.from({ length: n }, (_, i) => [i]);
  const alive = new Uint8Array(n).fill(1);
  let liveCount = n;

  while (liveCount > 1) {
    let best = Infinity;
    let bestA = -1;
    let bestB = -1;

    for (let a = 0; a < n; a++) {
      if (!alive[a]) continue;
      for (let b = a + 1; b < n; b++) {
        if (!alive[b]) continue;
        const d = distance[a * n + b];
        if (d < best) {
          best = d;
          bestA = a;
          bestB = b;
        }
      }
    }

    if (bestA === -1 || best >= granularity) break;

    const sizeA = members[bestA].length;
    const sizeB = members[bestB].length;

    for (let c = 0; c < n; c++) {
      if (!alive[c] || c === bestA || c === bestB) continue;
      const merged = (sizeA * distance[bestA * n + c] + sizeB * distance[bestB * n + c])
        / (sizeA + sizeB);
      distance[bestA * n + c] = merged;
      distance[c * n + bestA] = merged;
    }

    members[bestA] = [...members[bestA], ...members[bestB]];
    members[bestB] = [];
    alive[bestB] = 0;
    liveCount--;
  }

  const groups = [];
  for (let i = 0; i < n; i++) {
    if (alive[i]) groups.push(members[i]);
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
