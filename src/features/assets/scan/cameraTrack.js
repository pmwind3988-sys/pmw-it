/**
 * The things a phone camera can do beyond pointing at something.
 *
 * A store room is dark, a barcode on a pallet is far away, and a fixed-focus
 * frame of a sticker held at arm's length decodes to nothing. All three are
 * why a scan "never reads the barcode at all", and all three are settings on
 * the video track rather than anything the decoder can fix.
 *
 * Every one of them is OPTIONAL. Support varies by browser and by handset —
 * Chrome on Android has torch and zoom, Safari has neither at the time of
 * writing — so each is asked for, never required, and a phone that cannot do
 * it simply does not show the button.
 *
 * Untested on purpose, and kept as thin as it can be: none of it can run
 * without a camera in the room. What CAN be tested — which part of the frame
 * to read — lives in `cropRegion.js` beside it.
 */

/**
 * What to ask the camera for.
 *
 * 1080p rather than 720p because the bars of a barcode are the finest detail
 * the camera is ever pointed at, and `ideal` means a laptop webcam that cannot
 * manage it still opens instead of throwing.
 *
 * `continuous` focus is the important one. Without it many browsers hand back
 * a fixed-focus stream, and a label held closer than a metre is a blur that no
 * decoder will ever read.
 */
export const VIDEO_CONSTRAINTS = {
  facingMode: { ideal: 'environment' },
  width: { ideal: 1920 },
  height: { ideal: 1080 },
  // `advanced` is the browser's own "try these, skip what you cannot do" list,
  // which is exactly the contract wanted here.
  advanced: [{ focusMode: 'continuous' }],
};

const videoTrack = (stream) => stream?.getVideoTracks?.()[0] ?? null;

/**
 * What this particular phone turned out to support, in the shape the controls
 * want. Everything is absent rather than false where the browser does not
 * implement `getCapabilities` at all — which Safari does not.
 */
export function trackControls(stream) {
  const track = videoTrack(stream);
  if (!track?.getCapabilities) return { torch: false, zoom: null };

  let capabilities;
  try {
    capabilities = track.getCapabilities() ?? {};
  } catch {
    return { torch: false, zoom: null };
  }

  return {
    torch: capabilities.torch === true,
    // A range with no room to move is not a zoom control worth drawing.
    zoom: capabilities.zoom && capabilities.zoom.max > capabilities.zoom.min
      ? {
        min: capabilities.zoom.min,
        max: capabilities.zoom.max,
        step: capabilities.zoom.step || 0.1,
      }
      : null,
  };
}

/**
 * Applying one of them. Failure is swallowed deliberately: a phone refusing
 * the torch must not take the scan down with it, and there is nothing the
 * person holding it could do about the error anyway.
 */
async function apply(stream, constraint) {
  const track = videoTrack(stream);
  if (!track?.applyConstraints) return false;

  try {
    await track.applyConstraints({ advanced: [constraint] });
    return true;
  } catch {
    return false;
  }
}

export const setTorch = (stream, on) => apply(stream, { torch: Boolean(on) });

export const setZoom = (stream, zoom) => apply(stream, { zoom: Number(zoom) });

/**
 * Focus where the person just tapped. `pointsOfInterest` is in normalised
 * coordinates — 0 to 1 across the frame — which is what a tap position
 * divided by the element's size already is.
 */
export const focusAt = (stream, x, y) => apply(stream, {
  focusMode: 'single-shot',
  pointsOfInterest: [{ x, y }],
});
