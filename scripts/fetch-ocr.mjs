// Puts the text-recognition engine where the app can serve it itself.
//
// Same rule as the analysis model in `fetch-model.mjs`: everything the
// browser loads at RUNTIME must come from this application's own
// origin. Tesseract's default behaviour is the opposite -- it fetches
// its WebAssembly core and its language data from a public CDN the
// first time a worker starts, which is a request from every user's
// browser to a third party and an outright failure on a corporate
// network that blocks one.
//
// The engine and the worker are already in node_modules, so they are
// COPIED. Only the language data has to be downloaded, and it is
// hash-verified against `ocr-manifest.json` the same way the model is.
//
// Only the LSTM cores are copied. Tesseract ships a second set carrying
// the legacy engine as well, at roughly 600KB more each, and the reader
// asks for LSTM only -- so the legacy halves would be downloaded by
// every phone and used by none.
//
//   node scripts/fetch-ocr.mjs            verify against the manifest
//   node scripts/fetch-ocr.mjs --update   record a new hash
//
// Nothing here runs in the browser. `npm run build` runs it through the
// prebuild hook; a fresh clone that only runs `npm run dev` needs
// `npm run fetch:ocr` once, or the scan button reports that recognition
// is unavailable.

import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const MODULES = join(ROOT, 'node_modules');

/** Must match `OCR_BASE` in src/features/assets/scan/textReader.js. */
const OUT = join(ROOT, 'public', 'ocr');

/**
 * The worker script, and the three cores it may ask for. Which core a
 * phone loads is decided in the browser by feature detection, so all
 * three have to be on disk -- but only one is ever downloaded.
 *
 * The `.wasm.js` files carry the WebAssembly inside them; the bare
 * `.wasm` files beside them in node_modules are for bundlers and are
 * deliberately not copied.
 */
const COPIES = [
  ['tesseract.js/dist/worker.min.js', 'worker.min.js'],
  ['tesseract.js-core/tesseract-core-lstm.wasm.js', 'core/tesseract-core-lstm.wasm.js'],
  ['tesseract.js-core/tesseract-core-simd-lstm.wasm.js', 'core/tesseract-core-simd-lstm.wasm.js'],
  ['tesseract.js-core/tesseract-core-relaxedsimd-lstm.wasm.js', 'core/tesseract-core-relaxedsimd-lstm.wasm.js'],
];

/**
 * `4.0.0_best_int` is the LSTM-only build of the English data: the same
 * accuracy as the full file for printed labels, at a third of the size.
 */
const LANG_FILE = 'eng.traineddata.gz';
const LANG_URL = `https://cdn.jsdelivr.net/npm/@tesseract.js-data/eng/4.0.0_best_int/${LANG_FILE}`;

const MANIFEST = join(ROOT, 'scripts', 'ocr-manifest.json');

const update = process.argv.includes('--update');

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

async function loadManifest() {
  if (!existsSync(MANIFEST)) return {};
  return JSON.parse(await readFile(MANIFEST, 'utf8'));
}

async function copyEngine() {
  for (const [from, to] of COPIES) {
    const source = join(MODULES, from);
    if (!existsSync(source)) {
      throw new Error(`${from} is missing. Run npm install first.`);
    }

    const target = join(OUT, to);
    await mkdir(dirname(target), { recursive: true });
    await copyFile(source, target);
    process.stdout.write(`copied ${to}\n`);
  }
}

async function fetchLanguage(manifest) {
  const target = join(OUT, 'lang', LANG_FILE);

  if (existsSync(target) && !update && manifest[LANG_FILE]) {
    const existing = await readFile(target);
    if (sha256(existing) === manifest[LANG_FILE]) return;
  }

  process.stdout.write(`fetching ${LANG_FILE} ... `);
  const response = await fetch(LANG_URL);
  if (!response.ok) throw new Error(`${LANG_URL} -> HTTP ${response.status}`);
  const buffer = Buffer.from(await response.arrayBuffer());

  const hash = sha256(buffer);
  if (!update && manifest[LANG_FILE] && manifest[LANG_FILE] !== hash) {
    throw new Error(
      `${LANG_FILE} does not match the recorded hash. Re-run with --update only `
      + 'if you intend to accept different language data.',
    );
  }
  manifest[LANG_FILE] = hash;

  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, buffer);
  process.stdout.write(`${(buffer.length / 1e6).toFixed(1)}MB\n`);
}

const manifest = await loadManifest();
await copyEngine();
await fetchLanguage(manifest);
await writeFile(MANIFEST, `${JSON.stringify(manifest, null, 2)}\n`);
process.stdout.write('text recognition ready under public/ocr\n');
