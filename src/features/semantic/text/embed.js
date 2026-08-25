// The only module in this feature that is not a pure function.
//
// Everything here exists to keep one promise: no response text ever
// leaves the browser, and neither the model nor the runtime is fetched
// from anyone else's server.

import { pipeline, env } from '@huggingface/transformers';
export const MODEL_ID = 'Xenova/all-MiniLM-L6-v2';
export const BATCH_SIZE = 16;

// Served by this app, not by anyone else.
//
// `allowLocalModels` defaults to FALSE in the browser, so turning remote
// models off without turning local ones on disables both and the
// pipeline refuses to start. Verified against the running app: the error
// is "both local and remote models are disabled", which does at least
// say so, but only once someone opens the tab.
env.allowLocalModels = true;
env.allowRemoteModels = false;
env.localModelPath = '/models/';

// `wasmPaths` is deliberately NOT set.
//
// The runtime reaches its .wasm through `new URL(..., import.meta.url)`,
// which the bundler rewrites to a hashed asset on this origin. So the
// default already keeps the promise, and overriding it made the build
// carry the same 22.5MB file twice -- once bundled, once in public/ --
// while serving only the copy named here.
//
// If this ever has to be set again, it must name the ONE file rather
// than a directory prefix. A prefix makes the runtime treat its loader
// as external too and fetch it with a dynamic import(), which Vite's dev
// server refuses to do for anything in public/ -- so the prefix form
// works in a build and fails in dev. The file wanted is the asyncify
// build: transformers.js imports `onnxruntime-web/webgpu`, and that
// bundle's loader names `ort-wasm-simd-threaded.asyncify.wasm`.

// Loading the model dominates the first run, so the 0-40 band of the
// progress bar is reserved for it. Without its own progress the tab
// looks hung for the better part of ten seconds.
const LOAD_PCT = 40;
const EMBED_PCT = 45;

export function createEmbedder() {
  let extractor = null;
  let loading = null;

  function ready(onProgress) {
    if (extractor) return Promise.resolve(extractor);
    // One in-flight load, however many callers ask. Two concurrent
    // pipeline() calls would each fetch and compile the model.
    if (loading) return loading;

    loading = pipeline('feature-extraction', MODEL_ID, {
      dtype: 'q8',
      progress_callback: (event) => {
        if (event?.status !== 'progress') return;
        const fraction = (event.progress ?? 0) / 100;
        onProgress?.({ stage: 'Loading the model', pct: Math.round(fraction * LOAD_PCT) });
      },
    }).then((made) => {
      extractor = made;
      return made;
    });

    return loading;
  }

  async function embedAll(texts, { onProgress } = {}) {
    const list = texts ?? [];
    if (list.length === 0) return [];

    const run = await ready(onProgress);
    const out = [];

    for (let start = 0; start < list.length; start += BATCH_SIZE) {
      const batch = list.slice(start, start + BATCH_SIZE);
      // Mean pooling and L2 normalisation happen here so every consumer
      // can treat a vector as a direction and compare with plain cosine.
      const tensor = await run(batch, { pooling: 'mean', normalize: true });
      for (const row of tensor.tolist()) out.push(Float32Array.from(row));

      onProgress?.({
        stage: 'Understanding responses',
        pct: LOAD_PCT + Math.round(((start + batch.length) / list.length) * EMBED_PCT),
      });
    }

    return out;
  }

  return { embedAll };
}
