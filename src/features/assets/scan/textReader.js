import { scaledSize } from './capturePhoto.js';

/**
 * Turning a video frame into lines of text.
 *
 * The same shape as `detector.js`, and for the same reasons: some
 * browsers can do this themselves, everything above this file speaks one
 * interface — `read(source)` returning `[{ text, confidence }]` — and
 * that interface is what the tests fake.
 *
 * The engine is served from this application's own origin. Tesseract
 * would otherwise fetch its WebAssembly core and its English data from a
 * public CDN on first use, which is a request from every user's browser
 * to a third party and a plain failure on a network that blocks one.
 * `scripts/fetch-ocr.mjs` puts both under `public/ocr` at build time.
 */

export const READER_SOURCE = { NATIVE: 'native', TESSERACT: 'tesseract', NONE: 'none' };

/** Must match the output directory in scripts/fetch-ocr.mjs. */
export const OCR_BASE = '/ocr';

/**
 * What the frame is shrunk to before it is read.
 *
 * Recognition cost rises with the pixel count, and this runs several
 * times per scan on a phone. 1000px on the longest edge still resolves
 * label text held at arm's length — the same thing the item photograph
 * has to do at 1600px — while reading roughly twice as fast as the raw
 * 1280×720 camera frame.
 */
export const READ_EDGE = 1000;

/**
 * Sparse text. The default assumes a page of prose and tries to order
 * the label into paragraphs; a sticker is scattered short lines, and
 * telling the engine so is the difference between reading them and
 * merging two of them into one.
 */
const PSM_SPARSE_TEXT = '11';

/** LSTM only — the legacy engine is not among the files we ship. */
const OEM_LSTM_ONLY = 1;

function nativeReader() {
  if (typeof globalThis.TextDetector !== 'function') return null;
  try {
    return new globalThis.TextDetector();
  } catch {
    return null;
  }
}

/** A frame, shrunk onto a canvas the engine can take. */
export function toCanvas(source, maxEdge = READ_EDGE) {
  const width = source.videoWidth ?? source.naturalWidth ?? source.width;
  const height = source.videoHeight ?? source.naturalHeight ?? source.height;
  if (!width || !height) return null;

  const size = scaledSize(width, height, maxEdge);
  const canvas = document.createElement('canvas');
  canvas.width = size.width;
  canvas.height = size.height;
  canvas.getContext('2d').drawImage(source, 0, 0, size.width, size.height);
  return canvas;
}

/** Every line inside a recognition result, whatever depth it is nested at. */
function linesFromBlocks(blocks) {
  const lines = [];

  for (const block of blocks ?? []) {
    for (const paragraph of block.paragraphs ?? []) {
      for (const line of paragraph.lines ?? []) {
        lines.push({ text: String(line.text ?? ''), confidence: line.confidence });
      }
    }
  }

  return lines;
}

/**
 * Returns `{ read, source, terminate }`, or `{ read: null }` where
 * nothing on this browser can recognise text at all — which the scan
 * sheet turns into an explanation and the keyboard, rather than a camera
 * that silently finds nothing.
 */
export async function createTextReader() {
  const native = nativeReader();
  if (native) {
    return {
      source: READER_SOURCE.NATIVE,
      // The browser's own detector reports no confidence. `cleanLines`
      // treats a missing one as "no opinion" rather than as zero.
      read: async (frame) => (await native.detect(frame))
        .flatMap((found) => String(found.rawValue ?? '').split('\n'))
        .map((text) => ({ text })),
      terminate: () => {},
    };
  }

  try {
    const { createWorker } = await import('tesseract.js');
    const worker = await createWorker('eng', OEM_LSTM_ONLY, {
      workerPath: `${OCR_BASE}/worker.min.js`,
      corePath: `${OCR_BASE}/core`,
      langPath: `${OCR_BASE}/lang`,
      gzip: true,
      legacyCore: false,
      legacyLang: false,
    });
    await worker.setParameters({ tessedit_pageseg_mode: PSM_SPARSE_TEXT });

    return {
      source: READER_SOURCE.TESSERACT,
      read: async (frame) => {
        // `blocks` is what carries a confidence per line; the plain text
        // output is one string for the whole frame, which would make one
        // bad word discredit every line beside it.
        const { data } = await worker.recognize(frame, {}, { text: true, blocks: true });
        const lines = linesFromBlocks(data?.blocks);
        if (lines.length) return lines;

        return String(data?.text ?? '')
          .split('\n')
          .map((text) => ({ text, confidence: data?.confidence }));
      },
      terminate: () => { worker.terminate().catch(() => {}); },
    };
  } catch {
    return { source: READER_SOURCE.NONE, read: null, terminate: () => {} };
  }
}

/**
 * One frame's worth of lines, normalised.
 *
 * A reader that throws on a frame — a half-drawn video element, a frame
 * arriving during a resize — must not end the scanning loop. There will
 * be another frame; the right response to one bad one is to skip it.
 */
export async function readTextFrame(read, frame) {
  if (!read || !frame) return [];

  const canvas = toCanvas(frame);
  if (!canvas) return [];

  try {
    const found = await read(canvas);
    return (found ?? [])
      .map((entry) => ({
        text: String(entry.text ?? '').trim(),
        confidence: entry.confidence,
      }))
      .filter((entry) => entry.text);
  } catch {
    return [];
  }
}
