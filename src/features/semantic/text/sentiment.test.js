import { describe, it, expect } from 'vitest';
import { sentimentOf, sentimentTally, SENTIMENT } from './sentiment.js';

const label = (text) => sentimentOf(text).sentiment;

describe('sentimentOf', () => {
  it('reads a complaint as negative and a compliment as positive', () => {
    expect(label('It crashes every morning and the error is useless')).toBe(SENTIMENT.NEGATIVE);
    expect(label('Works well now, thanks — much quicker')).toBe(SENTIMENT.POSITIVE);
  });

  it('says nothing about an answer carrying neither', () => {
    expect(label('We use the shared drive for the monthly report')).toBe(SENTIMENT.NEUTRAL);
    expect(label('')).toBe(SENTIMENT.NEUTRAL);
    expect(label(null)).toBe(SENTIMENT.NEUTRAL);
  });

  /**
   * The whole reason the lexicon is matched on word boundaries. A substring
   * scan finds "not" inside "another", "bad" inside "badge" and "issue"
   * inside "issued" — and files three compliments as complaints.
   */
  it('does not find a word inside a longer one', () => {
    expect(label('Another badge was issued to the notice board')).toBe(SENTIMENT.NEUTRAL);
    expect(label('The hardware is fine and the software works well')).toBe(SENTIMENT.POSITIVE);
  });

  it('matches a contraction, which a word boundary alone would miss', () => {
    // `\bn't\b` never matches: there is no boundary before the n in "doesn't".
    expect(label("It doesn't sync and the report isn't right")).toBe(SENTIMENT.NEGATIVE);
  });

  it('leaves the IT nouns alone — a hard disk is not an opinion', () => {
    expect(label('Swapping the hard drive fixed it')).toBe(SENTIMENT.NEUTRAL);
    expect(label('It runs on the critical path server')).toBe(SENTIMENT.NEUTRAL);
  });

  it('calls a mixed answer neutral rather than picking a side', () => {
    // "fast" and "crashes" cancel. Calling this Positive would bury a fault;
    // calling it Negative would throw away what is working.
    expect(label('Fast, but it crashes')).toBe(SENTIMENT.NEUTRAL);
  });

  it('keeps the margin, so the worst of it can be sorted to the top', () => {
    const mild = sentimentOf('The report is missing a column');
    const furious = sentimentOf('Broken, useless, terrible — constant errors and downtime');

    expect(furious.score).toBeLessThan(mild.score);
  });

  it('caps one repetitive answer so it cannot outweigh four people', () => {
    const shouted = sentimentOf('bad bad bad bad bad bad bad bad bad bad');

    expect(shouted.score).toBe(-4);
  });
});

describe('sentimentTally', () => {
  it('counts how a set of fragments splits', () => {
    const fragments = [
      { sentiment: { sentiment: SENTIMENT.NEGATIVE } },
      { sentiment: { sentiment: SENTIMENT.NEGATIVE } },
      { sentiment: { sentiment: SENTIMENT.POSITIVE } },
      { },
    ];

    expect(sentimentTally(fragments)).toEqual({ Positive: 1, Neutral: 1, Negative: 2 });
  });

  it('is calm about nothing at all', () => {
    expect(sentimentTally()).toEqual({ Positive: 0, Neutral: 0, Negative: 0 });
  });
});
