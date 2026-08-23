import { describe, it, expect } from 'vitest';
import { scaledSize, MAX_EDGE } from './capturePhoto.js';

describe('scaledSize', () => {
  it('brings a phone photo down to the long edge', () => {
    expect(scaledSize(4032, 3024)).toEqual({ width: MAX_EDGE, height: 1200 });
  });

  it('works on a portrait photo too', () => {
    expect(scaledSize(3024, 4032)).toEqual({ width: 1200, height: MAX_EDGE });
  });

  /** Enlarging a small photo costs bytes and adds nothing. */
  it('leaves something already small alone', () => {
    expect(scaledSize(640, 480)).toEqual({ width: 640, height: 480 });
  });

  it('never rounds an edge away to nothing', () => {
    expect(scaledSize(4000, 1).height).toBe(1);
  });

  it('is calm about a source that reports no size yet', () => {
    expect(scaledSize(0, 0)).toEqual({ width: 0, height: 0 });
  });
});
