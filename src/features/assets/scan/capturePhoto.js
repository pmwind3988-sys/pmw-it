/**
 * Photographs, shrunk to a size a phone can hold a whole delivery of.
 *
 * A modern phone camera produces 4–8MB per shot. Thirty of those is a quarter
 * of a gigabyte sitting in IndexedDB waiting for a signal, which is how a
 * scanning session ends in a quota error instead of a saved delivery.
 *
 * 1600px on the longest edge at quality 0.72 lands around 200–350KB and still
 * reads a serial-number label held at arm's length, which is the only thing
 * these photographs have to do.
 */

export const MAX_EDGE = 1600;
export const QUALITY = 0.72;

export function scaledSize(width, height, maxEdge = MAX_EDGE) {
  const longest = Math.max(width, height);
  // Never scaled UP: enlarging a small photo costs bytes and adds nothing.
  if (!longest || longest <= maxEdge) return { width, height };

  const ratio = maxEdge / longest;
  return {
    width: Math.max(1, Math.round(width * ratio)),
    height: Math.max(1, Math.round(height * ratio)),
  };
}

/**
 * `source` is anything `drawImage` accepts — a video element for a frame
 * grabbed off the camera, an ImageBitmap for a file chosen from the gallery.
 */
export async function shrinkToBlob(source, { width, height, quality = QUALITY } = {}) {
  const size = scaledSize(
    width ?? source.videoWidth ?? source.naturalWidth ?? source.width,
    height ?? source.videoHeight ?? source.naturalHeight ?? source.height,
  );

  const canvas = document.createElement('canvas');
  canvas.width = size.width;
  canvas.height = size.height;
  canvas.getContext('2d').drawImage(source, 0, 0, size.width, size.height);

  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new Error('The photo could not be saved'))),
      'image/jpeg',
      quality,
    );
  });
}

/** A file the user picked, through the same shrink as a camera frame. */
export async function shrinkFile(file) {
  const bitmap = await createImageBitmap(file);
  try {
    return await shrinkToBlob(bitmap, { width: bitmap.width, height: bitmap.height });
  } finally {
    bitmap.close?.();
  }
}
