// The only module in this feature that is not a pure function.
//
// Everything here exists to keep one promise: no response text ever
// leaves the browser. Two settings do that, and the second is the one
// that is easy to miss -- without `wasmPaths` the ONNX runtime is
// fetched from a public CDN the first time anyone opens the tab, and
// nothing on screen says so, because the feature still works.

import { pipeline, env } from '@huggingface/transformers';

export const MODEL_ID = 'Xenova/all-MiniLM-L6-v2';
export const BATCH_SIZE = 16;

// Served by this app, not by anyone else.
env.allowRemoteModels = false;
env.localModelPath = '/models/';
env.backends.onnx.wasm.wasmPaths = '/ort/';

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
