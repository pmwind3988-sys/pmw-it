// The model's home -- text analysis spec, section 11.
//
// A worker of its own, not the studio worker. That one holds the parsed
// grid and has to answer a re-clean message the instant the user ticks a
// checkbox; giving it a 23MB model and multi-second inference sessions
// as well would make every clean wait behind an embedding run.
//
// The fragments and their vectors STAY here after an analysis, the same
// way the grid stays in the studio worker. A threshold change then costs
// one small message instead of re-embedding everything -- which is the
// difference between a control that feels live and one that does not.

import { analyze, rescore } from '../text/analysis.js';
import { createEmbedder } from '../text/embed.js';

const embedder = createEmbedder();

// Bucket vectors live as long as the worker does. Keyed by prompt text
// (see embedBuckets), so moving a slider or renaming a bucket reuses
// them and only an edit to what a bucket MEANS pays to embed again.
const bucketCache = new Map();

let current = null;

function report(progress) {
  self.postMessage({ type: 'progress', stage: progress.stage, pct: progress.pct });
}

const embedAll = (texts, options) => embedder.embedAll(texts, options);

async function handleAnalyze(msg) {
  report({ stage: 'Loading the model', pct: 2 });

  const raw = await analyze({
    texts: msg.texts,
    breadths: msg.breadths,
    buckets: msg.buckets,
    columnName: msg.columnName,
    settings: msg.settings,
    embedAll,
    bucketCache,
    onProgress: report,
  });

  current = {
    columnName: raw.columnName,
    fragments: raw.fragments,
    vectors: raw.vectors,
    noIssueRows: raw.noIssueRows,
  };

  self.postMessage({ type: 'analyzed', raw });
}

async function handleRescore(msg) {
  if (!current) throw new Error('There is nothing analysed to re-score.');

  report({ stage: 'Grouping', pct: 60 });
  const raw = await rescore({
    columnName: current.columnName,
    fragments: current.fragments,
    vectors: current.vectors,
    noIssueRows: current.noIssueRows,
    buckets: msg.buckets,
    settings: msg.settings,
    embedAll,
    bucketCache,
  });

  self.postMessage({ type: 'analyzed', raw });
}

self.onmessage = async (e) => {
  const msg = e.data ?? {};
  try {
    if (msg.type === 'analyze') await handleAnalyze(msg);
    else if (msg.type === 'rescore') await handleRescore(msg);
  } catch (err) {
    // Never leave the tab on a spinner. A model that will not load is
    // the most likely failure here, and the message has to say so, or
    // the user is left watching a progress bar that stopped.
    self.postMessage({
      type: 'error',
      message: err?.message || 'The analysis stopped unexpectedly.',
    });
  }
};
