// The whole pipeline, in order -- spec §6.
//
// The embedder arrives as an argument and is never imported here. That
// single choice is what keeps this file pure enough to test: every case
// in analysis.test.js runs against a four-dimensional stand-in, with no
// model, no WebAssembly and no network. The worker is the only place
// that supplies the real one.

import { splitIssues } from './splitIssues.js';
import { bucketPromptText, STARTER_BUCKETS } from './buckets.js';
import { severityOf } from './severity.js';
import { assignBuckets, meanVector, DEFAULT_THRESHOLD } from './similarity.js';
import { clusterVectors, DEFAULT_GRANULARITY } from './cluster.js';
import { labelCluster } from './labelCluster.js';
import { normalizeText } from './boilerplate.js';

// Below this, grouping is theatre. Two sentences do not have themes,
// and presenting a "theme" over four fragments invites the reader to
// draw a conclusion the data cannot support.
export const MIN_FRAGMENTS_FOR_THEMES = 5;

export function buildFragments(texts, breadths = []) {
  const rows = texts ?? [];

  // The mean is computed over the FRAGMENTS, not the raw answers, so the
  // length term in severity compares like with like.
  const perRow = rows.map((text) => splitIssues(text));
  const flat = perRow.flat();
  const meanLength = flat.length > 0
    ? flat.reduce((sum, t) => sum + t.length, 0) / flat.length
    : 80;

  const fragments = [];
  for (let row = 0; row < perRow.length; row++) {
    perRow[row].forEach((text, index) => {
      fragments.push({
        id: `${row}:${index}`,
        row,
        index,
        text,
        severity: severityOf(text, { meanLength, breadth: breadths[row] ?? 0 }),
      });
    });
  }
  return fragments;
}

function noIssueRowsOf(texts, fragments) {
  const withIssues = new Set(fragments.map((f) => f.row));
  const rows = [];
  for (let row = 0; row < (texts ?? []).length; row++) {
    if (!withIssues.has(row)) rows.push(row);
  }
  return rows;
}

// One averaged vector per bucket, built from its description and hints.
async function embedBuckets(buckets, embedAll) {
  const prompts = buckets.map(bucketPromptText);
  const flat = prompts.flat();
  if (flat.length === 0) return [];

  const vectors = await embedAll(flat);

  let cursor = 0;
  return buckets.map((bucket, i) => {
    const slice = vectors.slice(cursor, cursor + prompts[i].length);
    cursor += prompts[i].length;
    return { id: bucket.id, vector: meanVector(slice) };
  });
}

function discoverThemes(fragments, vectors, granularity) {
  if (fragments.length < MIN_FRAGMENTS_FOR_THEMES) {
    return { themes: [], oneOffIds: fragments.map((f) => f.id), themeByFragment: new Map() };
  }

  const { clusters, oneOffs } = clusterVectors(vectors, granularity);
  const allTexts = fragments.map((f) => f.text);
  const themeByFragment = new Map();

  const themes = clusters.map((members, i) => {
    const id = `theme_${i}`;
    const fragmentIds = members.map((index) => fragments[index].id);
    for (const fragmentId of fragmentIds) themeByFragment.set(fragmentId, id);
    return {
      id,
      name: labelCluster(members.map((index) => allTexts[index]), allTexts),
      fragmentIds,
    };
  });

  return { themes, oneOffIds: oneOffs.map((index) => fragments[index].id), themeByFragment };
}

function assemble({
  columnName, buckets, fragments, assignments,
  themes, oneOffIds, themeByFragment, noIssueRows, settings,
}) {
  return {
    columnName,
    settings,
    buckets,
    fragments: fragments.map((fragment, i) => ({
      ...fragment,
      bucketId: assignments[i].bucketId,
      score: assignments[i].score,
      themeId: themeByFragment.get(fragment.id) ?? null,
    })),
    themes,
    oneOffIds,
    noIssueRows,
  };
}

export async function analyze({
  texts, breadths = [], buckets = STARTER_BUCKETS, columnName = '',
  settings = {}, embedAll, onProgress = () => {},
}) {
  const threshold = settings.threshold ?? DEFAULT_THRESHOLD;
  const granularity = settings.granularity ?? DEFAULT_GRANULARITY;
  const resolved = { threshold, granularity };

  onProgress({ stage: 'Reading responses', pct: 42 });
  const fragments = buildFragments(texts, breadths);
  const noIssueRows = noIssueRowsOf(texts, fragments);

  if (fragments.length === 0) {
    return {
      columnName,
      settings: resolved,
      buckets,
      fragments: [],
      themes: [],
      oneOffIds: [],
      noIssueRows,
      vectors: [],
    };
  }

  onProgress({ stage: 'Understanding responses', pct: 45 });
  const vectors = await embedAll(fragments.map((f) => normalizeText(f.text)), { onProgress });

  const bucketVectors = await embedBuckets(buckets, embedAll);
  const assignments = assignBuckets(vectors, bucketVectors, threshold);

  onProgress({ stage: 'Grouping', pct: 85 });
  const { themes, oneOffIds, themeByFragment } = discoverThemes(fragments, vectors, granularity);

  onProgress({ stage: 'Ranking', pct: 95 });
  return {
    ...assemble({
      columnName,
      buckets,
      fragments,
      assignments,
      themes,
      oneOffIds,
      themeByFragment,
      noIssueRows,
      settings: resolved,
    }),
    // Kept so a threshold or granularity change never re-embeds. This is
    // what the sub-second settings budget in spec §16 rests on.
    vectors,
  };
}

/**
 * Re-file and re-group WITHOUT re-embedding the fragments.
 *
 * Only the bucket descriptions are embedded here -- a dozen short
 * strings. Anything that re-embeds fragments on a settings change is a
 * regression against the performance budget, not an optimisation
 * opportunity missed.
 */
export async function rescore({
  columnName = '', fragments, vectors, buckets = STARTER_BUCKETS,
  settings = {}, noIssueRows = [], embedAll,
}) {
  const threshold = settings.threshold ?? DEFAULT_THRESHOLD;
  const granularity = settings.granularity ?? DEFAULT_GRANULARITY;
  const resolved = { threshold, granularity };

  if ((fragments ?? []).length === 0) {
    return {
      columnName,
      settings: resolved,
      buckets,
      fragments: [],
      themes: [],
      oneOffIds: [],
      noIssueRows,
      vectors: [],
    };
  }

  const bucketVectors = await embedBuckets(buckets, embedAll);
  const assignments = assignBuckets(vectors, bucketVectors, threshold);
  const { themes, oneOffIds, themeByFragment } = discoverThemes(fragments, vectors, granularity);

  return {
    ...assemble({
      columnName,
      buckets,
      fragments,
      assignments,
      themes,
      oneOffIds,
      themeByFragment,
      noIssueRows,
      settings: resolved,
    }),
    vectors,
  };
}
