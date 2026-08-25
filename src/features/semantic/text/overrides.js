// The user's corrections, applied over the model's answer -- spec §8.
//
// The raw analysis is never mutated. Corrections live in their own
// record and are re-applied to whatever the model most recently
// produced, which is what makes three things true at once:
//
//   * re-running the model, moving the threshold, or editing a bucket
//     description never destroys a correction;
//   * "reset to what the model said" is discarding one object;
//   * an override naming a fragment that no longer exists is dropped
//     rather than corrupting the screen -- a re-import with different
//     data is a normal thing to do.

import { rankIssues } from './rankIssues.js';
import { UNSORTED_ID, UNSORTED_LABEL } from './buckets.js';

export const EMPTY_OVERRIDES = {
  retags: {},
  noise: [],
  themeNames: {},
  themeMerges: {},
  pinned: [],
  suppressed: [],
};

// Follows a chain of merges to whichever theme is still standing, and
// refuses to loop if two themes were somehow merged into each other.
function resolveTheme(themeId, merges) {
  let current = themeId;
  const seen = new Set();
  while (merges[current] && !seen.has(current)) {
    seen.add(current);
    current = merges[current];
  }
  return current;
}

function summarise(fragmentIds, byId) {
  const respondents = new Set();
  let severityTotal = 0;
  let counted = 0;

  for (const id of fragmentIds) {
    const fragment = byId.get(id);
    if (!fragment || fragment.noise) continue;
    respondents.add(fragment.row);
    severityTotal += fragment.severity ?? 0;
    counted++;
  }

  return {
    count: counted,
    respondents: respondents.size,
    meanSeverity: counted > 0 ? severityTotal / counted : 0,
  };
}

export function applyOverrides(raw, overrides = EMPTY_OVERRIDES) {
  const {
    retags = {}, noise = [], themeNames = {}, themeMerges = {},
    pinned = [], suppressed = [],
  } = overrides ?? {};

  const noiseSet = new Set(noise);
  const bucketIds = new Set([...(raw?.buckets ?? []).map((b) => b.id), UNSORTED_ID]);

  const fragments = (raw?.fragments ?? []).map((fragment) => {
    // A retag naming a bucket that has since been deleted falls back to
    // the model's answer rather than to a bucket nothing can render.
    const retag = retags[fragment.id];
    const bucketId = retag && bucketIds.has(retag) ? retag : fragment.bucketId;
    return {
      ...fragment,
      bucketId,
      themeId: resolveTheme(fragment.themeId, themeMerges),
      noise: noiseSet.has(fragment.id),
    };
  });

  const byId = new Map(fragments.map((f) => [f.id, f]));

  const byBucket = new Map();
  for (const fragment of fragments) {
    if (!byBucket.has(fragment.bucketId)) byBucket.set(fragment.bucketId, []);
    byBucket.get(fragment.bucketId).push(fragment.id);
  }

  const definitions = [
    ...(raw?.buckets ?? []),
    // Always present, even when empty: the pile where the model declined
    // to guess is information, and an absent Unsorted reads as "nothing
    // was ambiguous".
    { id: UNSORTED_ID, label: UNSORTED_LABEL, description: '', hints: [] },
  ];

  const buckets = definitions.map((definition) => {
    const fragmentIds = byBucket.get(definition.id) ?? [];
    return { ...definition, fragmentIds, ...summarise(fragmentIds, byId) };
  });

  const byTheme = new Map();
  for (const fragment of fragments) {
    if (!fragment.themeId) continue;
    if (!byTheme.has(fragment.themeId)) byTheme.set(fragment.themeId, []);
    byTheme.get(fragment.themeId).push(fragment.id);
  }

  const themes = (raw?.themes ?? [])
    // A theme merged into another one no longer exists on its own.
    .filter((theme) => !themeMerges[theme.id])
    .map((theme) => {
      const fragmentIds = byTheme.get(theme.id) ?? [];
      return {
        ...theme,
        name: themeNames[theme.id] ?? theme.name,
        fragmentIds,
        ...summarise(fragmentIds, byId),
      };
    });

  const priority = rankIssues(
    [
      ...buckets
        .filter((b) => b.id !== UNSORTED_ID && b.count > 0)
        .map((b) => ({
          kind: 'bucket',
          id: b.id,
          label: b.label,
          respondents: b.respondents,
          count: b.count,
          meanSeverity: b.meanSeverity,
        })),
      ...themes
        .filter((t) => t.count > 0)
        .map((t) => ({
          kind: 'theme',
          id: t.id,
          label: t.name,
          respondents: t.respondents,
          count: t.count,
          meanSeverity: t.meanSeverity,
        })),
    ],
    { pinned, suppressed },
  );

  return { ...raw, fragments, buckets, themes, priority };
}
