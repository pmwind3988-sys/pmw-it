// What to call a theme nobody named -- spec §6.6.
//
// c-TF-IDF: a term earns its place by being frequent INSIDE the cluster
// and rare outside it. Plain frequency would name every theme after the
// survey's own subject -- "data · process · work" on all of them --
// which tells the reader nothing about which theme they are looking at.
//
// The result is a starting point. Renaming a theme is a first-class
// action in the UI precisely because four words are a label, not a
// sentence.

import { normalizeText } from './boilerplate.js';

const STOPWORDS = new Set([
  'the', 'a', 'an', 'and', 'or', 'but', 'if', 'then', 'so', 'because', 'as',
  'of', 'in', 'on', 'at', 'to', 'for', 'with', 'from', 'by', 'into', 'about',
  'is', 'are', 'was', 'were', 'be', 'been', 'being', 'am',
  'do', 'does', 'did', 'have', 'has', 'had', 'having',
  'it', 'its', 'this', 'that', 'these', 'those', 'there', 'here',
  'we', 'our', 'us', 'i', 'my', 'me', 'you', 'your', 'they', 'them', 'their',
  'he', 'she', 'his', 'her', 'which', 'who', 'whom', 'what', 'when', 'where',
  'not', 'no', 'nor', 'can', 'cannot', 'could', 'will', 'would', 'should',
  'may', 'might', 'must', 'need', 'needs', 'very', 'more', 'most', 'much',
  'many', 'some', 'any', 'all', 'every', 'each', 'other', 'than', 'also',
  'up', 'out', 'down', 'over', 'under', 'again', 'only', 'just', 'still',
]);

const MIN_TERM_LENGTH = 3;

export const SEPARATOR = ' · ';
export const UNNAMED = 'Unnamed theme';

function tokenize(text) {
  return normalizeText(text)
    .toLowerCase()
    .split(/[^a-z0-9-]+/)
    .filter((token) => token.length >= MIN_TERM_LENGTH && !STOPWORDS.has(token));
}

export function labelCluster(memberTexts, allTexts, termCount = 4) {
  const members = memberTexts ?? [];
  if (members.length === 0) return UNNAMED;

  const corpus = allTexts ?? [];
  const total = Math.max(1, corpus.length);

  // How many fragments in the WHOLE corpus contain each term.
  const documentFrequency = new Map();
  for (const text of corpus) {
    for (const token of new Set(tokenize(text))) {
      documentFrequency.set(token, (documentFrequency.get(token) ?? 0) + 1);
    }
  }

  const termFrequency = new Map();
  for (const text of members) {
    for (const token of tokenize(text)) {
      termFrequency.set(token, (termFrequency.get(token) ?? 0) + 1);
    }
  }

  const scored = [];
  for (const [term, tf] of termFrequency) {
    const df = documentFrequency.get(term) ?? 1;
    // A term in every fragment scores log(1) = 0 and drops out, which is
    // exactly the "data · process · work" problem this exists to solve.
    const idf = Math.log(total / df);
    if (idf <= 0) continue;
    scored.push({ term, score: tf * idf });
  }

  if (scored.length === 0) return UNNAMED;

  scored.sort((a, b) => b.score - a.score || a.term.localeCompare(b.term));
  return scored.slice(0, termCount).map((s) => s.term).join(SEPARATOR);
}
