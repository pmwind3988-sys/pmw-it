/**
 * Which part of the camera frame the aiming box is actually over.
 *
 * The single biggest reason a barcode is never read: the whole 1280x720 frame
 * is handed to the decoder, so a small sticker held at arm's length is a
 * handful of pixels wide and resolves to nothing. Decoding the aiming box
 * alone, at the camera's own resolution, gives the same sticker several times
 * the pixels across the bars.
 *
 * The geometry is not decoration. The viewfinder is `object-fit: cover`, so a
 * landscape camera in a portrait window has most of its width cropped away
 * before anybody sees it — reading the whole frame decodes barcodes that were
 * never on screen and misses the one being pointed at. This works out what
 * cover is showing and takes the box out of THAT.
 *
 * Pure, and the only part of the scanner that can be tested without a camera
 * in the room.
 */

/**
 * The aiming box, as a fraction in from each edge. Must match `.as-reticle`
 * and `.as-sheet-frame` in `assets.css`, or the crop and the rectangle people
 * aim with are two different rectangles.
 */
export const RETICLE_INSET = { x: 0.1, y: 0.18 };

const clamp = (value, low, high) => Math.min(Math.max(value, low), high);

export function cropRegion({
  sourceWidth,
  sourceHeight,
  viewWidth,
  viewHeight,
  inset = RETICLE_INSET,
}) {
  // A video element that has not painted a frame yet reports zero for all of
  // these. Returning a crop of nothing would end the decode loop.
  if (!sourceWidth || !sourceHeight || !viewWidth || !viewHeight) return null;

  // `cover` scales by whichever axis needs the most, and centres the overflow.
  const scale = Math.max(viewWidth / sourceWidth, viewHeight / sourceHeight);
  const shownWidth = viewWidth / scale;
  const shownHeight = viewHeight / scale;
  const originX = (sourceWidth - shownWidth) / 2;
  const originY = (sourceHeight - shownHeight) / 2;

  const insetX = clamp(inset.x ?? 0, 0, 0.49);
  const insetY = clamp(inset.y ?? 0, 0, 0.49);

  const x = Math.round(originX + shownWidth * insetX);
  const y = Math.round(originY + shownHeight * insetY);
  const width = Math.round(shownWidth * (1 - insetX * 2));
  const height = Math.round(shownHeight * (1 - insetY * 2));

  return {
    x: clamp(x, 0, sourceWidth),
    y: clamp(y, 0, sourceHeight),
    width: clamp(width, 1, sourceWidth - clamp(x, 0, sourceWidth)),
    height: clamp(height, 1, sourceHeight - clamp(y, 0, sourceHeight)),
  };
}

/**
 * That region of the frame, drawn onto a canvas at its own resolution — no
 * shrinking, because shrinking is the problem this exists to solve.
 */
export function cropToCanvas(video, inset = RETICLE_INSET) {
  const region = cropRegion({
    sourceWidth: video?.videoWidth,
    sourceHeight: video?.videoHeight,
    viewWidth: video?.clientWidth,
    viewHeight: video?.clientHeight,
    inset,
  });
  if (!region) return null;

  const canvas = document.createElement('canvas');
  canvas.width = region.width;
  canvas.height = region.height;
  canvas.getContext('2d').drawImage(
    video,
    region.x, region.y, region.width, region.height,
    0, 0, region.width, region.height,
  );
  return canvas;
}
