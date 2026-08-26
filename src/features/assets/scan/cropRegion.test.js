import { describe, it, expect } from 'vitest';
import { cropRegion, RETICLE_INSET } from './cropRegion.js';

const inset = { x: 0.1, y: 0.2 };

describe('cropRegion', () => {
  /**
   * The straightforward case: the picture on screen is the whole frame,
   * shrunk. The aiming box is a tenth in from each side, so the crop is too.
   */
  it('maps the aiming box onto the frame when nothing is cropped away', () => {
    expect(cropRegion({
      sourceWidth: 1280, sourceHeight: 720, viewWidth: 640, viewHeight: 360, inset,
    })).toEqual({
      x: 128, y: 144, width: 1024, height: 432,
    });
  });

  /**
   * The case that matters, and the one that makes the aiming box mean
   * something. The video is `object-fit: cover`, so a landscape camera in a
   * portrait window has most of its width cropped away before anybody sees it.
   * Reading the whole frame decodes a barcode that was never on screen, and
   * misses the one the person is pointing at.
   */
  it('follows what cover actually shows when the shapes disagree', () => {
    expect(cropRegion({
      sourceWidth: 1280, sourceHeight: 720, viewWidth: 360, viewHeight: 640, inset,
    })).toEqual({
      x: 478, y: 144, width: 324, height: 432,
    });
  });

  it('never runs off the edge of the frame', () => {
    const region = cropRegion({
      sourceWidth: 1280,
      sourceHeight: 720,
      viewWidth: 360,
      viewHeight: 640,
      inset: { x: -1, y: -1 },
    });

    expect(region.x).toBeGreaterThanOrEqual(0);
    expect(region.y).toBeGreaterThanOrEqual(0);
    expect(region.x + region.width).toBeLessThanOrEqual(1280);
    expect(region.y + region.height).toBeLessThanOrEqual(720);
  });

  /**
   * A video element that has not painted a frame yet reports zero for
   * everything, and a crop of nothing would end the decode loop.
   */
  it('says nothing rather than guessing before the camera has a picture', () => {
    expect(cropRegion({
      sourceWidth: 0, sourceHeight: 0, viewWidth: 360, viewHeight: 640, inset,
    })).toBe(null);
    expect(cropRegion({
      sourceWidth: 1280, sourceHeight: 720, viewWidth: 0, viewHeight: 0, inset,
    })).toBe(null);
  });

  /** The default has to match the box drawn in the stylesheet, or the crop
   *  and the rectangle people aim with are two different rectangles. */
  it('defaults to the box the stylesheet draws', () => {
    expect(RETICLE_INSET).toEqual({ x: 0.1, y: 0.18 });

    const region = cropRegion({
      sourceWidth: 1280, sourceHeight: 720, viewWidth: 1280, viewHeight: 720,
    });

    expect(region).toEqual({
      x: 128, y: 130, width: 1024, height: 461,
    });
  });
});
