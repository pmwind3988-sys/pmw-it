// The Semantic Analysis context object and its stage machine, apart from
// the component that provides it.
//
// Two separate project rules land on this file. A module that exports a
// component must export nothing else or it drops out of Fast Refresh and
// fails lint -- which is why `useSemantic` lives in its own file and why
// the context object and these constants live here rather than beside
// the provider.

import { createContext } from 'react';

export const SemanticContext = createContext(null);

// Where the user is. There are only three places to be, and two of them
// are transient: a sheet is dropped, it is read, and from then on the
// user is looking at their answers. Nothing here parks them on a
// profile or a cleaning checklist -- every question those screens used
// to ask is now answered by the app.
export const STAGES = ['idle', 'parsing', 'dashboard', 'text'];

export const IDLE_STATE = {
  stage: 'idle',
  fileName: '',
  sheets: [],
  activeSheet: '',
  headerIndex: -1,
  headerCandidates: [],
  grid: null,
  profile: null,
  // Which columns were parked as form bookkeeping. Kept so the brief
  // card can list them and put them back in one click.
  overrides: {},
  // The cleaning checklist, and the typed dataset that results from
  // applying it. It is proposed and applied without asking: on a forms
  // export the answer to every question the old review screen asked was
  // the one the proposal already had. `dataset` is null until the worker
  // returns the first clean.
  plan: [],
  dataset: null,
  cleaning: false,
  // The charts, the filters the user set from the filter bar, and the
  // chart click currently in force.
  //
  // `selection` and `globalFilters` are kept apart because they behave
  // differently -- a selection spares the chart it came from, so its
  // other bars stay clickable; a global filter applies everywhere.
  tiles: [],
  globalFilters: [],
  selection: null,
  error: '',
  progress: { stage: '', pct: 0 },
  // --- semantic analysis ------------------------------------------------
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
  // --- what the app decided for itself ----------------------------------
  // The subject it read out of the file name, the bookkeeping columns it
  // parked, and which written-answer column it read. Null until a file
  // is parsed.
  //
  // `autoSeed` is the one-shot flag that survives the gap between "the
  // sheet is parsed" and "the cleaned dataset came back from the
  // worker", which is the moment the starter charts can actually be
  // chosen. `autoAnalysed` and `autoCharted` are the same idea one step
  // later: read the writing once, and chart what the reading found once.
  brief: null,
  autoSeed: false,
  autoAnalysed: false,
  autoCharted: false,
};

// Chart widths on the 12-column grid.
export const TILE_SIZES = { S: 3, M: 6, L: 9, XL: 12 };
export const SIZE_ORDER = ['S', 'M', 'L', 'XL'];
