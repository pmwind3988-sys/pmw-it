// One written answer -> the separate issues inside it (spec §6.2).
//
// People do not answer "describe your challenge" with one challenge.
// They write a paragraph containing three, and counting the paragraph
// once undercounts every problem in it but the first.
//
// The splitting is rules-based on purpose. The model that follows can
// tell whether two fragments mean the same thing; it cannot tell where
// one thought ends -- and a sentence boundary is a perfectly good
// answer to that question in this data, where respondents write in
// clean sentences.

import { stripLabelPrefix, isNonAnswer } from './boilerplate.js';

// A fragment shorter than this is a split that went wrong -- an
// abbreviation, an initial, a stray "etc." -- and is folded back into
// the fragment before it rather than becoming an issue of its own.
const MIN_FRAGMENT_LENGTH = 25;

// A ceiling, not a target. A pathological answer must not produce a
// hundred rows the user has to read.
export const MAX_FRAGMENTS = 12;

const BULLET_RE = /^\s*(?:[-–—•*]|\d+[.)])\s+/;

// A sentence ends at .!? followed by space and something that starts a
// new sentence. The lookbehind excludes the common abbreviations that
// otherwise cut a sentence in half.
const SENTENCE_SPLIT_RE = /(?<!\b(?:e\.g|i\.e|etc|vs|no|dr|mr|ms|mrs)\.)(?<=[.!?])\s+(?=[A-Z0-9"'(])/;

export function splitIssues(text) {
  // Newlines are a boundary, so the raw value is split on them BEFORE
  // normalising -- normalizeText collapses them into ordinary spaces and
  // the structure would be gone.
  const lines = String(text ?? '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map(stripLabelPrefix)
    .filter((line) => line !== '');

  const pieces = [];
  for (const line of lines) {
    const withoutBullet = line.replace(BULLET_RE, '').trim();
    if (withoutBullet === '') continue;
    for (const sentence of withoutBullet.split(SENTENCE_SPLIT_RE)) {
      const trimmed = sentence.trim();
      if (trimmed !== '') pieces.push(trimmed);
    }
  }

  // Fold a too-short piece into the one before it. Nothing to fold into
  // means it is the first piece, and it stands on its own.
  const merged = [];
  for (const piece of pieces) {
    if (piece.length < MIN_FRAGMENT_LENGTH && merged.length > 0) {
      merged[merged.length - 1] = `${merged[merged.length - 1]} ${piece}`;
      continue;
    }
    merged.push(piece);
  }

  return merged.filter((piece) => !isNonAnswer(piece)).slice(0, MAX_FRAGMENTS);
}
