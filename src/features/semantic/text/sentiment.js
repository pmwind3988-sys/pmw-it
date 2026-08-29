// Whether an answer reads as a complaint, a compliment, or neither.
//
// Distinct from `severity.js`, which measures how STRONGLY something is
// expressed. The two answer different questions and a reader wants both:
// severity says "this person is exercised about it", sentiment says which
// direction they are pointing. A long emphatic "the new scanner has saved
// us an hour a day" scores high on severity and is not a problem at all,
// and until now nothing in the pipeline could tell it from a complaint.
//
// A word count, deliberately. The model in this feature is good at whether
// two fragments MEAN the same thing and is not a sentiment classifier; a
// lexicon is honest about being a rough signal, is inspectable when it gets
// one wrong, and costs nothing on top of an inference run that is already
// the slow part.
//
// It is a SIGNAL, not a verdict -- the same caveat severity carries, for the
// same reason. Nothing here knows that "it works, but only if you restart
// it twice" is a complaint.

import { normalizeText } from './boilerplate.js';

/**
 * Matched on WORD BOUNDARIES, never as substrings.
 *
 * This is the whole reason the list is compiled rather than scanned with
 * `includes`: in an IT survey "hardware" contains "hard", "another" and
 * "notice" contain "not", and "badge" contains "bad". A substring scan reads
 * "the hardware is fine" as two negative words and one positive, and files a
 * compliment as a complaint. Multi-word entries are matched as phrases.
 */
const NEGATIVE_WORDS = [
  'not', 'never', 'cannot', "can't", 'cant', 'unable', 'fail', 'fails', 'failed', 'failure',
  'slow', 'slowly', 'delay', 'delays', 'delayed', 'wait', 'waiting', 'stuck', 'blocked', 'block',
  'broken', 'breaks', 'crash', 'crashes', 'error', 'errors', 'bug', 'bugs', 'issue', 'issues',
  'problem', 'problems', 'difficult', 'complicated', 'complex', 'confusing', 'confused',
  'frustrating', 'frustrated', 'annoying', 'annoyed', 'painful', 'terrible', 'awful', 'bad',
  'worse', 'worst', 'useless', 'unusable', 'unreliable', 'unstable', 'outdated', 'lost',
  'waste', 'wasted', 'wasting', 'tedious', 'manual', 'repetitive', 'duplicate', 'missing',
  'lack', 'lacking', 'poor', 'insufficient', 'too many', 'too much', 'too long', 'nightmare',
  'chaos', 'urgent', 'blocker', 'downtime', 'jam', 'freeze', 'frozen', 'complain',
];

/**
 * Two words are deliberately absent, and both for the same reason: they are
 * ordinary IT NOUNS here rather than opinions. "hard" is a hard drive and
 * "critical" is a critical system or a critical path -- neither says anybody
 * is unhappy, and both are common enough in this data to tilt whole themes.
 * Word boundaries fix "hardware" containing "hard"; they cannot fix "hard
 * disk", where the word really is "hard".
 */

const POSITIVE_WORDS = [
  'good', 'great', 'excellent', 'fine', 'fast', 'quick', 'quickly', 'easy', 'easier', 'easily',
  'simple', 'smooth', 'smoothly', 'reliable', 'stable', 'happy', 'satisfied', 'helpful',
  'thanks', 'thank you', 'appreciate', 'appreciated', 'perfect', 'nice', 'love', 'improved',
  'improvement', 'better', 'best', 'works well', 'no issues', 'no problem',
  'well done', 'efficient', 'convenient', 'clear', 'useful', 'solved', 'resolved',
];

/**
 * Matched as the END of a word rather than as a word of its own: "doesn't",
 * "isn't", "won't", "couldn't". These cannot live in the list above, because
 * a left word boundary makes them unmatchable — there is no boundary between
 * the s and the n in "doesn't", so `n't` never fires. One entry covers
 * every English negative contraction, which is why the list is this short.
 */
const NEGATIVE_SUFFIXES = ["n't"];

export const SENTIMENT = { POSITIVE: 'Positive', NEUTRAL: 'Neutral', NEGATIVE: 'Negative' };

const escape = (term) => term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * `\b` is wrong at an apostrophe: `\bn't\b` never matches, because there is
 * no word boundary before `n` in "doesn't". Terms that do not start with a
 * word character are anchored on the left by whatever precedes them instead.
 */
function compile(terms) {
  return terms.map((term) => {
    const left = /^\w/.test(term) ? '\\b' : '';
    const right = /\w$/.test(term) ? '\\b' : '';
    return new RegExp(`${left}${escape(term)}${right}`, 'g');
  });
}

/** Anchored on the right only, so it can attach to the word before it. */
function compileSuffixes(terms) {
  return terms.map((term) => new RegExp(`${escape(term)}\\b`, 'g'));
}

const NEGATIVE_RE = [...compile(NEGATIVE_WORDS), ...compileSuffixes(NEGATIVE_SUFFIXES)];
const POSITIVE_RE = compile(POSITIVE_WORDS);

// Four of a kind is as far as one side is allowed to run. Without a ceiling
// a single long answer that repeats itself outweighs four different people
// -- the same trap `severity.js` caps `INTENSITY_SATURATION` against.
const SATURATION = 4;

function countMatches(lower, patterns) {
  let hits = 0;
  for (const pattern of patterns) {
    pattern.lastIndex = 0;
    while (pattern.exec(lower) !== null) {
      hits += 1;
      if (hits >= SATURATION) return SATURATION;
    }
  }
  return hits;
}

/**
 * `{ sentiment, score }` -- the label, and the margin it was decided by.
 *
 * The score is kept because the label alone throws away how close the call
 * was, and a reader sorting for the worst of it wants -4 above -1. Positive
 * is a genuinely useful answer here rather than an afterthought: a survey
 * asking what is wrong still collects "the new laptops are excellent", and
 * counting that as an issue is how a priority list ends up with a theme
 * nobody needs to act on.
 */
export function sentimentOf(text) {
  const normalized = normalizeText(text);
  if (normalized === '') return { sentiment: SENTIMENT.NEUTRAL, score: 0 };

  const lower = normalized.toLowerCase();
  const negative = countMatches(lower, NEGATIVE_RE);
  const positive = countMatches(lower, POSITIVE_RE);
  const score = positive - negative;

  if (score < 0) return { sentiment: SENTIMENT.NEGATIVE, score };
  if (score > 0) return { sentiment: SENTIMENT.POSITIVE, score };
  // A tie is not a middle. An answer carrying both is mixed, and calling it
  // Neutral is the honest reading of "fast, but it crashes every morning".
  return { sentiment: SENTIMENT.NEUTRAL, score: 0 };
}

/** How a set of fragments splits, for a heading that says so at a glance. */
export function sentimentTally(fragments = []) {
  const tally = { Positive: 0, Neutral: 0, Negative: 0 };
  for (const fragment of fragments) {
    const label = fragment?.sentiment?.sentiment ?? SENTIMENT.NEUTRAL;
    if (label in tally) tally[label] += 1;
  }
  return tally;
}
