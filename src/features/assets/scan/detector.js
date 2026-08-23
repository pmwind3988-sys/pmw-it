/**
 * Turning a video frame into barcodes.
 *
 * Chrome on Android decodes them natively and fast, several in one frame.
 * Safari has no such thing, so a WebAssembly decoder stands in — imported
 * dynamically, so an Android phone never downloads the megabyte it does not
 * need.
 *
 * Everything above this file speaks one interface: `detect(source)` returning
 * `[{ rawValue, format }]`. That is also what the tests fake, which is the
 * point — no camera is under test, but the decisions taken about what comes
 * out of one are.
 */

/**
 * The formats worth looking for. Restricting the list matters: an unrestricted
 * detector spends its frame budget hunting formats no equipment label uses,
 * and the scan visibly slows down.
 *
 * `code_128` and `code_39` carry serials; `data_matrix` is what small parts
 * are marked with; `ean_13` / `upc_a` are the retail codes on a sealed box;
 * `qr_code` because some manufacturers now put the whole label in one.
 */
export const FORMATS = [
  'code_128', 'code_39', 'code_93', 'codabar',
  'data_matrix', 'qr_code', 'pdf417', 'itf',
  'ean_13', 'ean_8', 'upc_a', 'upc_e',
];

export const DETECTOR_SOURCE = { NATIVE: 'native', PONYFILL: 'ponyfill', NONE: 'none' };

function nativeDetector() {
  if (typeof globalThis.BarcodeDetector !== 'function') return null;
  try {
    return new globalThis.BarcodeDetector({ formats: FORMATS });
  } catch {
    // A browser can carry the constructor and still refuse this format list.
    // Falling through to the ponyfill is better than refusing to scan.
    return null;
  }
}

/**
 * Returns `{ detect, source }`, or `{ detect: null }` where nothing can decode
 * at all — which the scan screen turns into an explanation and a link to
 * manual entry, rather than a camera that silently never finds anything.
 */
export async function createDetector() {
  const native = nativeDetector();
  if (native) {
    return {
      source: DETECTOR_SOURCE.NATIVE,
      detect: (frame) => native.detect(frame),
    };
  }

  try {
    const { BarcodeDetector } = await import('barcode-detector/pure');
    const fallback = new BarcodeDetector({ formats: FORMATS });
    return {
      source: DETECTOR_SOURCE.PONYFILL,
      detect: (frame) => fallback.detect(frame),
    };
  } catch {
    return { source: DETECTOR_SOURCE.NONE, detect: null };
  }
}

/**
 * One frame's worth of codes, normalised to what the session expects.
 *
 * A decoder that throws on a frame — a half-drawn video element, a frame
 * arriving during a resize — must not end the scanning loop. There will be
 * another frame in sixteen milliseconds; the right response to one bad one is
 * to skip it.
 */
export async function readFrame(detect, frame) {
  if (!detect || !frame) return [];

  try {
    const found = await detect(frame);
    return (found ?? [])
      .map((entry) => ({
        rawValue: String(entry.rawValue ?? '').trim(),
        format: String(entry.format ?? ''),
      }))
      .filter((entry) => entry.rawValue);
  } catch {
    return [];
  }
}
