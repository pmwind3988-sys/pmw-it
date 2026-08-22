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
export const STAGES = ['idle', 'parsing', 'profiled', 'cleaning', 'canvas'];

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
  error: '',
  progress: { stage: '', pct: 0 },
};
