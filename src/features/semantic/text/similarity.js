// Fragment -> bucket, or an honest refusal -- spec §6.4.
//
// The refusal is the point. Every fragment has a nearest bucket, and
// filing it there regardless produces a screen where everything is
// categorised and some of it is wrong -- with nothing marking which.
// Below the threshold a fragment goes to Unsorted, which is a visible
// pile the user can act on.

import { UNSORTED_ID } from './buckets.js';

// Calibrated against real survey text during implementation; exposed in
// the UI because no single value is right for every survey.
export const DEFAULT_THRESHOLD = 0.3;

export function cosine(a, b) {
  let dot = 0;
  let na = 0;
  let nb = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    dot += a[i] * b[i];
    na += a[i] * a[i];
    nb += b[i] * b[i];
  }
  // A zero vector has no direction, so it is not similar to anything --
  // 0, never NaN, which would poison every comparison downstream.
  if (na === 0 || nb === 0) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

export function meanVector(vectors) {
  const list = vectors ?? [];
  if (list.length === 0) return Float32Array.from([]);

  const out = new Float32Array(list[0].length);
  for (const vector of list) {
    for (let i = 0; i < out.length; i++) out[i] += vector[i] ?? 0;
  }

  let norm = 0;
  for (let i = 0; i < out.length; i++) {
    out[i] /= list.length;
    norm += out[i] * out[i];
  }
  norm = Math.sqrt(norm);
  if (norm > 0) {
    for (let i = 0; i < out.length; i++) out[i] /= norm;
  }
  return out;
}

export function assignBuckets(fragmentVectors, bucketVectors, threshold = DEFAULT_THRESHOLD) {
  const buckets = bucketVectors ?? [];

  return (fragmentVectors ?? []).map((vector) => {
    let bestId = UNSORTED_ID;
    let bestScore = -Infinity;

    for (const bucket of buckets) {
      const score = cosine(vector, bucket.vector);
      if (score > bestScore) {
        bestScore = score;
        bestId = bucket.id;
      }
    }

    if (buckets.length === 0 || bestScore < threshold) {
      // The score is still reported for Unsorted rows -- it is what the
      // "lower the threshold" prompt reads to say how close they were.
      return { bucketId: UNSORTED_ID, score: buckets.length === 0 ? 0 : bestScore };
    }
    return { bucketId: bestId, score: bestScore };
  });
}
