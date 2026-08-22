// Puts the analysis model where the app can serve it itself.
//
// The model must be served from this application's own origin -- see
// the text analysis spec, section 2. Fetching it from a public host at
// RUNTIME would send a request from every user's browser to a third
// party and would fail outright on a corporate network that blocks it.
// Fetching it at BUILD time and serving it ourselves has neither
// problem.
//
// The files are gitignored rather than committed, which keeps ~25MB out
// of every clone at the cost of a build-time network dependency. If
// that ever becomes unacceptable, commit public/models and public/ort
// and delete the prebuild hook; nothing else changes.
//
//   node scripts/fetch-model.mjs            verify against the manifest
//   node scripts/fetch-model.mjs --update   record new hashes

import { createHash } from 'node:crypto';
import {
  mkdir, readFile, writeFile, copyFile, readdir,
} from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const MODEL_ID = 'Xenova/all-MiniLM-L6-v2';
const REVISION = 'main';

const MODEL_FILES = [
  'config.json',
  'tokenizer.json',
  'tokenizer_config.json',
  'onnx/model_quantized.onnx',
];

const MANIFEST = join(ROOT, 'scripts', 'model-manifest.json');
const MODEL_DIR = join(ROOT, 'public', 'models', MODEL_ID);
const ORT_DIR = join(ROOT, 'public', 'ort');

// The runtime ships inside the installed package, so the runtime half of
// the promise needs no network at all. It lives in onnxruntime-web, NOT
// in the transformers dist directory -- that one carries only the .mjs
// loader and none of the .wasm binaries it loads.
const ORT_SOURCE = join(ROOT, 'node_modules', 'onnxruntime-web', 'dist');

const update = process.argv.includes('--update');

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

async function loadManifest() {
  if (!existsSync(MANIFEST)) return {};
  return JSON.parse(await readFile(MANIFEST, 'utf8'));
}

async function fetchModelFiles(manifest) {
  for (const file of MODEL_FILES) {
    const target = join(MODEL_DIR, file);

    if (existsSync(target) && !update && manifest[file]) {
      const existing = await readFile(target);
      if (sha256(existing) === manifest[file]) continue;
    }

    const url = `https://huggingface.co/${MODEL_ID}/resolve/${REVISION}/${file}`;
    process.stdout.write(`fetching ${file} ... `);
    const response = await fetch(url);
    if (!response.ok) throw new Error(`${url} -> HTTP ${response.status}`);
    const buffer = Buffer.from(await response.arrayBuffer());

    const hash = sha256(buffer);
    if (!update && manifest[file] && manifest[file] !== hash) {
      throw new Error(
        `${file} does not match the recorded hash. Re-run with --update only `
        + 'if you intend to accept a different model.',
      );
    }
    manifest[file] = hash;

    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, buffer);
    process.stdout.write(`${(buffer.length / 1e6).toFixed(1)}MB\n`);
  }
}

async function copyOnnxRuntime() {
  if (!existsSync(ORT_SOURCE)) {
    throw new Error('onnxruntime-web is not installed; run npm install first.');
  }
  await mkdir(ORT_DIR, { recursive: true });

  let copied = 0;
  for (const entry of await readdir(ORT_SOURCE)) {
    if (!entry.endsWith('.wasm') && !entry.endsWith('.mjs')) continue;
    await copyFile(join(ORT_SOURCE, entry), join(ORT_DIR, entry));
    copied++;
  }
  if (copied === 0) throw new Error(`No runtime files found in ${ORT_SOURCE}.`);
  process.stdout.write(`copied ${copied} ONNX runtime files\n`);
}

const manifest = await loadManifest();
await fetchModelFiles(manifest);
await copyOnnxRuntime();
await writeFile(MANIFEST, `${JSON.stringify(manifest, null, 2)}\n`);
process.stdout.write('model ready under public/\n');
