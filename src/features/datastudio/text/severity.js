// How strongly an issue is expressed -- spec §6.7.
//
// This is a SIGNAL, not a judgement, and the UI says so. Nothing here
// knows whether a problem is important; it measures four things that
// correlate with someone being frustrated enough to write about it, and
// the priority ranking then weighs that behind how many people raised
// it at all.
//
// The breadth term is the only one not inferred from prose: it comes
// from the structured multi-select column. Somebody who ticked seven
// challenges is telling us something the wording cannot.

import { normalizeText } from './boilerplate.js';

export const INTENSITY_TERMS = [
  'time-consuming', 'time consuming', 'prone to error', 'error-prone', 'error prone',
  'manual', 'manually', 'repetitive', 'repeatedly', 'delay', 'delays', 'delayed',
  'missed', 'overlooked', 'bottleneck', 'tedious', 'frustrating', 'frustration',
  'duplicate', 'duplication', 'rework', 'chase', 'chasing', 'constantly', 'always',
  'difficult', 'cannot', 'unable', 'no way to', 'waste', 'wasted', 'slow', 'stuck',
];

const WEIGHT_INTENSITY = 0.5;
const WEIGHT_LENGTH = 0.2;
const WEIGHT_BREADTH = 0.2;
const WEIGHT_EMPHASIS = 0.1;

// Four matches is as angry as this measure gets. Without a ceiling a
// single long answer that repeats itself outscores four different
// people, which is the ordering the ranking exists to prevent.
const INTENSITY_SATURATION = 4;

function countIntensity(lower) {
  let matches = 0;
  for (const term of INTENSITY_TERMS) {
    if (lower.includes(term)) matches++;
    if (matches >= INTENSITY_SATURATION) break;
  }
  return matches;
}

function emphasisOf(text) {
  const bangs = (text.match(/!/g) ?? []).length;
  const shouted = (text.match(/\b[A-Z]{3,}\b/g) ?? []).length;
  return Math.min(1, (bangs + shouted) / 3);
}

export function severityOf(text, { meanLength = 80, breadth = 0 } = {}) {
  const normalized = normalizeText(text);
  if (normalized === '') return 0;

  const lower = normalized.toLowerCase();

  const intensity = countIntensity(lower) / INTENSITY_SATURATION;
  // Relative to the corpus, capped: twice the average length is as much
  // as length is allowed to say.
  const length = Math.min(1, normalized.length / (Math.max(1, meanLength) * 2));
  const spread = Math.min(1, Math.max(0, breadth));
  const emphasis = emphasisOf(normalized);

  const score = intensity * WEIGHT_INTENSITY
    + length * WEIGHT_LENGTH
    + spread * WEIGHT_BREADTH
    + emphasis * WEIGHT_EMPHASIS;

  return Math.min(1, Math.max(0, score));
}
