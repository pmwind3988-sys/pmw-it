import { describe, it, expect } from 'vitest';
import { STARTER_BUCKETS, UNSORTED_ID, bucketPromptText } from './buckets.js';

describe('STARTER_BUCKETS', () => {
  it('gives every bucket a unique id', () => {
    const ids = STARTER_BUCKETS.map((b) => b.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('never ships a bucket called unsorted', () => {
    // Unsorted is where the model declines to guess. A real bucket with
    // that id would make a refusal indistinguishable from a match.
    expect(STARTER_BUCKETS.some((b) => b.id === UNSORTED_ID)).toBe(false);
  });

  it('describes every bucket in a sentence, not a label', () => {
    for (const bucket of STARTER_BUCKETS) {
      expect(bucket.description.length).toBeGreaterThan(30);
      expect(bucket.hints.length).toBeGreaterThan(0);
    }
  });
});

describe('bucketPromptText', () => {
  it('embeds the description and the hints, not the name', () => {
    const bucket = { id: 'x', label: 'SAP', description: 'Problems with SAP transactions.', hints: ['posting errors'] };
    expect(bucketPromptText(bucket)).toEqual(['Problems with SAP transactions.', 'posting errors']);
  });

  it('falls back to the label when someone clears the description', () => {
    expect(bucketPromptText({ id: 'x', label: 'SAP', description: '', hints: [] })).toEqual(['SAP']);
  });
});
