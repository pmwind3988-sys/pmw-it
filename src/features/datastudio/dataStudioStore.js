// The Data Studio context object and its stage machine, apart from the
// component that provides it.
//
// Two separate project rules land on this file. A module that exports a
// component must export nothing else or it drops out of Fast Refresh and
// fails lint -- which is why `useDataStudio` lives in its own file and
// why the context object and these constants live here rather than
// beside the provider.

import { createContext } from 'react';

export const DataStudioContext = createContext(null);

// Where the user is in the import -> clean -> chart journey. The page
// renders one screen per stage, so this is the only thing deciding what
// is on screen.
export const STAGES = ['idle', 'parsing', 'profiled', 'cleaning', 'canvas', 'text'];

export const IDLE_STATE = {
  stage: 'idle',
  fileName: '',
  sheets: [],
  activeSheet: '',
  headerIndex: -1,
  headerCandidates: [],
  grid: null,
  profile: null,
  overrides: {},
  // The proposed cleaning checklist, and the typed dataset that results
  // from applying its ticked steps. `dataset` is null until the worker
  // returns the first clean.
  plan: [],
  dataset: null,
  cleaning: false,
  // Per-column user decisions that the inference could not make alone:
  // which reading an ambiguous date column gets, and whether a datetime
  // column was stored as UTC.
  dateOrders: {},
  zones: {},
  // The canvas: the tiles on it, the filters the user set from the
  // filter bar, and the cross-filter click currently in force.
  //
  // `selection` and `globalFilters` are kept apart because they behave
  // differently -- a selection spares the tile it came from (spec
  // §10.3), a global filter applies everywhere.
  tiles: [],
  globalFilters: [],
  selection: null,
  editingTileId: null,
  // The id this import is saved under, once it has been saved. Null
  // means "not persisted yet", which is what the Save button reads.
  datasetId: null,
  storageFull: false,
  error: '',
  progress: { stage: '', pct: 0 },
  // --- text analysis ---------------------------------------------------
  // `rawAnalysis` is what the model said; `textOverrides` is what the
  // user said about it; `analysis` is the two combined and is the only
  // one anything renders. Keeping the first two apart is what lets a
  // re-score leave hand corrections standing.
  textColumns: [],
  textColumnName: '',
  rawAnalysis: null,
  analysis: null,
  buckets: [],
  textOverrides: null,
  textSettings: { threshold: 0.3, granularity: 0.45 },
  analysing: false,
  textProgress: { stage: '', pct: 0 },
  textError: '',
  // --- autopilot -------------------------------------------------------
  // What the app worked out from the file name and did about it: the
  // subject it read, the bookkeeping columns it parked, and which
  // written-answer column it went to read. Null until a file is parsed,
  // and null forever for a dataset reopened from the library -- the
  // brief describes an import, not a saved dashboard.
  //
  // `autoSeed` is the one-shot flag that survives the gap between "the
  // sheet is parsed" and "the cleaned dataset came back from the
  // worker", which is the moment the starter charts can actually be
  // chosen. Without it a second clean would re-seed the canvas over
  // tiles the user had already edited.
  brief: null,
  autoSeed: false,
  autoAnalysed: false,
};

// Tile widths on the 12-column canvas grid (spec §10.4).
export const TILE_SIZES = { S: 3, M: 6, L: 9, XL: 12 };
export const SIZE_ORDER = ['S', 'M', 'L', 'XL'];
