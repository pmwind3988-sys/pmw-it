import { describe, it, expect } from 'vitest';
import { gridToSend } from './gridSync.js';

const grid = { headers: ['A'], rows: [[1]] };

describe('gridToSend', () => {
  it('sends nothing when the worker already holds this grid', () => {
    expect(gridToSend(grid, grid)).toBeUndefined();
  });

  it('sends the grid when the main thread has replaced it', () => {
    // The case that matters: adding the text analysis appends five
    // columns and builds a NEW grid. Left unsent, the worker cleans the
    // previous one and every tile charting the analysis reports a
    // column that is plainly there.
    const withAnalysis = { headers: ['A', 'Severity'], rows: [[1, 40]] };
    expect(gridToSend(withAnalysis, grid)).toBe(withAnalysis);
  });

  it('sends the first grid the worker has never been told about', () => {
    expect(gridToSend(grid, null)).toBe(grid);
  });

  it('compares by identity, not by content', () => {
    // A deep comparison of 100k rows on every checkbox tick would cost
    // more than the message it is trying to save.
    expect(gridToSend({ ...grid }, grid)).toBeTruthy();
  });

  it('has nothing to send before anything is parsed', () => {
    expect(gridToSend(null, null)).toBeUndefined();
  });
});
