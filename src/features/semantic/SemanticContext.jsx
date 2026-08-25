import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { SemanticContext, IDLE_STATE } from './semanticStore.js';
import { proposeCleanPlan } from './clean/proposeCleanPlan.js';
import { suggestCharts } from './suggest/suggestCharts.js';
import { SIZE_ORDER } from './semanticStore.js';
import { detectTextColumns } from './text/detectTextColumns.js';
import { STARTER_BUCKETS } from './text/buckets.js';
import { applyOverrides, EMPTY_OVERRIDES } from './text/overrides.js';
import { withAnalysisCharted } from './text/chartAnalysis.js';
import { gridToSend } from './worker/gridSync.js';
import { planAutopilot } from './intent/planAutopilot.js';
import { hideColumns, unhideColumns } from './intent/hideColumns.js';

/**
 * Owns the workers and the stage machine for the whole section.
 *
 * Nothing in here writes to disk, to SharePoint or to the network. A
 * dropped file is read in this tab, charted in this tab, and gone when
 * the tab closes — the only way anything leaves is an export the user
 * asks for. That is not an implementation detail, it is the promise the
 * screen makes, and it is why there is no dataset library, no save
 * button and no IndexedDB anywhere below this file.
 *
 * The bytes of the imported file are kept in a ref, not in state:
 * changing sheet or correcting the header row means re-parsing the same
 * file, and asking the user to drop it again for that would be absurd.
 * They are in a ref rather than state because nothing renders from them
 * and a few MB in a state setter would re-render the tree for no
 * visible reason.
 */
export function SemanticProvider({ children }) {
  const [state, setState] = useState(IDLE_STATE);
  const workerRef = useRef(null);
  const textWorkerRef = useRef(null);
  const bytesRef = useRef(null);
  // The file name, beside the bytes rather than read from state. The
  // worker's message handler is installed once and closes over the
  // state of that first render, so reading `state.fileName` inside it
  // would give the autopilot an empty title on every import.
  const fileNameRef = useRef('');
  // Monotonic id per clean request. A result whose id is not the latest
  // is stale and applying it would show the answer to a question the
  // data has already moved past.
  const cleanIdRef = useRef(0);
  // The grid the WORKER is holding. It keeps the parsed grid so a
  // re-clean stays a small message, which means the main thread has to
  // notice when it has replaced that grid -- otherwise the worker
  // cleans the previous sheet and the difference shows up only as
  // charts reporting a column that is plainly on screen. See
  // `worker/gridSync.js`.
  const workerGridRef = useRef(null);

  useEffect(() => {
    // Vite's native worker syntax -- no plugin, and the worker's own
    // imports (SheetJS included) are bundled into a separate chunk that
    // the main thread never downloads.
    const worker = new Worker(new URL('./worker/parse.worker.js', import.meta.url), {
      type: 'module',
    });

    worker.onmessage = (e) => {
      const msg = e.data ?? {};

      if (msg.type === 'progress') {
        setState((s) => ({ ...s, progress: { stage: msg.stage, pct: msg.pct } }));
        return;
      }

      if (msg.type === 'parsed') {
        // The worker parsed it and is holding it; re-sending it on the
        // first clean would ship the whole sheet back for nothing.
        workerGridRef.current = msg.grid;
        // Everything decided about this file is decided here, once,
        // from the parse result -- not spread across the effects below.
        // The provider's job after this is only to carry the plan out.
        const textColumns = detectTextColumns(msg.profile, { rows: msg.grid.rows });
        const brief = planAutopilot({
          fileName: fileNameRef.current,
          sheetName: msg.activeSheet ?? '',
          profile: msg.profile,
          textColumns,
        });
        const profile = hideColumns(msg.profile, brief.hidden);

        setState((s) => ({
          ...s,
          stage: 'dashboard',
          brief,
          autoSeed: true,
          autoAnalysed: false,
          autoCharted: false,
          sheets: msg.sheets,
          activeSheet: msg.activeSheet,
          headerIndex: msg.headerIndex,
          headerCandidates: msg.headerCandidates ?? [],
          grid: msg.grid,
          profile,
          // The bookkeeping columns go in as ordinary overrides, so the
          // brief can list them and put them back in one click.
          overrides: Object.fromEntries(
            brief.hidden.map((c) => [c.name, { role: 'ignored' }]),
          ),
          plan: proposeCleanPlan(profile, msg.grid),
          dataset: null,
          // A different sheet is different writing, so an analysis of
          // the previous one describes nothing that is still on screen.
          textColumns,
          textColumnName: '',
          tiles: [],
          globalFilters: [],
          selection: null,
          rawAnalysis: null,
          analysis: null,
          textOverrides: null,
          textError: '',
          error: '',
          progress: { stage: '', pct: 100 },
        }));
        return;
      }

      if (msg.type === 'cleaned') {
        setState((s) => {
          if (msg.requestId !== cleanIdRef.current) return s;
          const next = { ...s, dataset: msg.dataset, cleaning: false };
          // The starter charts need the typed dataset, which only
          // exists now -- suggesting them at parse time would mean
          // suggesting them without the values a scatter is scored on.
          // One shot only: every later clean leaves the charts alone.
          if (!s.autoSeed) return next;
          return {
            ...next,
            autoSeed: false,
            tiles: s.tiles.length > 0
              ? s.tiles
              : suggestCharts(
                s.profile,
                msg.dataset,
                s.brief?.focus ?? [],
                s.textColumns.map((c) => c.name),
              ),
          };
        });
        return;
      }

      if (msg.type === 'error') {
        // Never leave the UI on a spinner -- go back to a screen the
        // user can act on, carrying the reason.
        setState((s) => ({
          ...s,
          stage: s.stage === 'parsing' ? 'idle' : s.stage,
          cleaning: false,
          error: msg.message,
          progress: { stage: '', pct: 0 },
        }));
      }
    };

    worker.onerror = (event) => {
      setState((s) => ({
        ...s,
        stage: 'idle',
        cleaning: false,
        error: event?.message || 'The import worker stopped unexpectedly.',
        progress: { stage: '', pct: 0 },
      }));
    };

    workerRef.current = worker;
    return () => {
      worker.terminate();
      workerRef.current = null;
      // The text worker is created lazily below, so it may never exist.
      textWorkerRef.current?.terminate();
      textWorkerRef.current = null;
    };
  }, []);

  /**
   * The analysis worker, created on first use rather than on mount.
   *
   * Constructing it eagerly would pull the model chunk into every visit
   * to this screen, including the ones that never drop a file.
   */
  const textWorker = useCallback(() => {
    if (textWorkerRef.current) return textWorkerRef.current;

    const worker = new Worker(new URL('./worker/text.worker.js', import.meta.url), {
      type: 'module',
    });

    worker.onmessage = (e) => {
      const msg = e.data ?? {};

      if (msg.type === 'progress') {
        setState((s) => ({ ...s, textProgress: { stage: msg.stage, pct: msg.pct } }));
        return;
      }

      if (msg.type === 'analyzed') {
        setState((s) => {
          const overrides = s.textOverrides ?? EMPTY_OVERRIDES;
          const next = {
            ...s,
            rawAnalysis: msg.raw,
            analysis: applyOverrides(msg.raw, overrides),
            textOverrides: overrides,
            analysing: false,
            textError: '',
            textProgress: { stage: '', pct: 100 },
          };
          // Charted here rather than from an effect watching `analysis`,
          // so that a re-score triggered by a slider does not silently
          // rebuild the charts underneath the user. The first reading
          // charts itself; every later one is charted on request.
          return s.autoCharted ? next : withAnalysisCharted(next);
        });
        return;
      }

      if (msg.type === 'error') {
        setState((s) => ({
          ...s,
          analysing: false,
          textError: msg.message,
          textProgress: { stage: '', pct: 0 },
        }));
      }
    };

    worker.onerror = (event) => {
      setState((s) => ({
        ...s,
        analysing: false,
        textError: event?.message || 'The analysis worker stopped unexpectedly.',
        textProgress: { stage: '', pct: 0 },
      }));
    };

    textWorkerRef.current = worker;
    return worker;
  }, []);

  const post = useCallback((sheetName, headerIndex) => {
    const bytes = bytesRef.current;
    if (!bytes || !workerRef.current) return;
    setState((s) => ({
      ...s, stage: 'parsing', error: '', progress: { stage: 'Reading workbook', pct: 5 },
    }));
    // A copy per post: the structured clone leaves ours intact, but only
    // because we never hand it over in a transfer list. Keep it that way.
    workerRef.current.postMessage({
      type: 'parse', arrayBuffer: bytes, sheetName, headerIndex,
    });
  }, []);

  // Asks the worker to re-apply the plan. The grid never crosses the
  // boundary again -- the worker kept it -- so this costs a small plan
  // message each way however large the sheet is.
  const requestClean = useCallback((profile, plan, grid) => {
    if (!workerRef.current || !profile || !plan) return;
    cleanIdRef.current += 1;
    setState((s) => ({ ...s, cleaning: true }));
    workerRef.current.postMessage({
      type: 'clean', profile, plan, grid, requestId: cleanIdRef.current,
    });
  }, []);

  const importFile = useCallback(async (file) => {
    if (!file) return;
    setState((s) => ({
      ...s,
      stage: 'parsing',
      fileName: file.name,
      error: '',
      progress: { stage: 'Reading file', pct: 2 },
    }));
    fileNameRef.current = file.name;
    try {
      bytesRef.current = await file.arrayBuffer();
    } catch {
      setState((s) => ({ ...s, stage: 'idle', error: `Could not read "${file.name}".` }));
      return;
    }
    post(undefined, undefined);
  }, [post]);

  const selectSheet = useCallback((name) => post(name, undefined), [post]);

  // Re-parsing for a header change re-detects nothing else: the sheet
  // stays put and only the split between header and body moves.
  const setHeaderIndex = useCallback(
    (index) => setState((s) => {
      post(s.activeSheet, index);
      return s;
    }),
    [post],
  );

  /**
   * Put the parked bookkeeping columns back among the charts.
   *
   * The brief keeps its list afterwards rather than clearing it, so the
   * card can go on saying which columns were parked and offer to park
   * them again. Only the roles and the overrides move.
   */
  const showHiddenColumns = useCallback(() => setState((s) => {
    const hidden = s.brief?.hidden ?? [];
    if (hidden.length === 0) return s;

    const overrides = { ...s.overrides };
    for (const column of hidden) delete overrides[column.name];
    const profile = unhideColumns(s.profile, hidden);

    return {
      ...s,
      profile,
      overrides,
      plan: proposeCleanPlan(profile, s.grid),
      brief: { ...s.brief, hiddenShown: true },
    };
  }), []);

  const hideAdminColumns = useCallback(() => setState((s) => {
    const hidden = s.brief?.hidden ?? [];
    if (hidden.length === 0) return s;

    const profile = hideColumns(s.profile, hidden);
    return {
      ...s,
      profile,
      overrides: {
        ...s.overrides,
        ...Object.fromEntries(hidden.map((c) => [c.name, { role: 'ignored' }])),
      },
      plan: proposeCleanPlan(profile, s.grid),
      brief: { ...s.brief, hiddenShown: false },
    };
  }), []);

  const dismissBrief = useCallback(
    () => setState((s) => (s.brief ? { ...s, brief: { ...s.brief, dismissed: true } } : s)), [],
  );

  // --- the charts -------------------------------------------------------

  const removeTile = useCallback((id) => setState((s) => ({
    ...s,
    tiles: s.tiles.filter((t) => t.id !== id),
    // A selection that came from a chart that no longer exists would
    // filter everything with nothing on screen explaining why.
    selection: s.selection?.sourceTileId === id ? null : s.selection,
  })), []);

  const cycleTileSize = useCallback((id) => setState((s) => ({
    ...s,
    tiles: s.tiles.map((t) => (t.id === id
      ? { ...t, size: SIZE_ORDER[(SIZE_ORDER.indexOf(t.size ?? 'M') + 1) % SIZE_ORDER.length] }
      : t)),
  })), []);

  // --- filters and cross-filter selection --------------------------------

  const addFilter = useCallback((filter) => setState((s) => ({
    ...s,
    // One filter per column: a second filter on the same column would
    // AND with the first and silently produce an empty screen.
    globalFilters: [...s.globalFilters.filter((f) => f.column !== filter.column), filter],
  })), []);

  const removeFilter = useCallback((column) => setState((s) => ({
    ...s,
    globalFilters: s.globalFilters.filter((f) => f.column !== column),
  })), []);

  const clearFilters = useCallback(
    () => setState((s) => ({ ...s, globalFilters: [] })), [],
  );

  const clearSelection = useCallback(
    () => setState((s) => (s.selection ? { ...s, selection: null } : s)), [],
  );

  /**
   * A tap on a bar, a slice or a point. `additive` is a shift-click.
   *
   * Tapping the value that is already the whole selection clears it --
   * the same gesture in and out, so there is always a way back without
   * hunting for a clear button.
   */
  const selectMark = useCallback(({ tileId, column, value, additive = false }) => {
    setState((s) => {
      if (!column || value === undefined || value === null) return s;

      const current = s.selection;
      const sameSource = current
        && current.sourceTileId === tileId
        && current.column === column;

      if (!sameSource) {
        return { ...s, selection: { sourceTileId: tileId, column, values: [value] } };
      }

      const values = additive
        ? (current.values.includes(value)
          ? current.values.filter((v) => v !== value)
          : [...current.values, value])
        : (current.values.length === 1 && current.values[0] === value ? [] : [value]);

      return { ...s, selection: values.length === 0 ? null : { ...current, values } };
    });
  }, []);

  // --- semantic analysis ------------------------------------------------

  // How many multi-select options each respondent picked, normalised
  // against the most anyone picked. This is the one severity input that
  // is measured rather than inferred, and it comes from the structured
  // column, not from the prose.
  const breadthsOf = useCallback((grid, profile) => {
    const multi = (profile?.columns ?? []).find((c) => c.type === 'multi');
    if (!multi) return grid.rows.map(() => 0);

    const separator = multi.separator ?? ';';
    const counts = grid.rows.map((row) => String(row?.[multi.index] ?? '')
      .split(separator)
      .map((part) => part.trim())
      .filter(Boolean).length);

    const most = Math.max(1, ...counts);
    return counts.map((n) => n / most);
  }, []);

  /**
   * Read a written-answer column.
   *
   * `navigate` is false when the import starts this by itself. The
   * analysis then runs in the background while the user reads the
   * charts that are already up, and the brief card shows how far it has
   * got -- being thrown onto a progress bar would be the feature
   * getting in its own way.
   */
  const startAnalysis = useCallback((columnName, { navigate = true } = {}) => setState((s) => {
    const column = s.textColumns.find((c) => c.name === columnName) ?? s.textColumns[0];
    if (!column || !s.grid) return s;

    const buckets = s.buckets.length > 0 ? s.buckets : STARTER_BUCKETS;
    textWorker().postMessage({
      type: 'analyze',
      columnName: column.name,
      texts: s.grid.rows.map((row) => row?.[column.index] ?? ''),
      breadths: breadthsOf(s.grid, s.profile),
      buckets,
      settings: s.textSettings,
    });

    return {
      ...s,
      stage: navigate ? 'text' : s.stage,
      textColumnName: column.name,
      buckets,
      // Latched here rather than by the effect below, so that a reading
      // started by hand also counts: whichever route got here, the
      // sheet has now had its answers read once and nothing should
      // queue a second pass over the same column.
      autoAnalysed: true,
      analysing: true,
      textError: '',
      textProgress: { stage: 'Loading the model', pct: 1 },
    };
  }), [textWorker, breadthsOf]);

  // Re-file against cached vectors. Never re-embeds the fragments --
  // that is what keeps a slider live rather than a five-second wait.
  const rescoreNow = useCallback((buckets, settings) => {
    textWorkerRef.current?.postMessage({ type: 'rescore', buckets, settings });
  }, []);

  const setTextSetting = useCallback((key, value) => setState((s) => {
    const textSettings = { ...s.textSettings, [key]: value };
    if (s.rawAnalysis) rescoreNow(s.buckets, textSettings);
    return { ...s, textSettings, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  const updateBucket = useCallback((id, patch) => setState((s) => {
    const buckets = s.buckets.map((b) => (b.id === id ? { ...b, ...patch } : b));
    // Only a description or hint change alters what matches. Renaming a
    // category is presentation, and re-running the model for it would
    // make typing in the name field stutter for no result.
    const rematches = patch.description !== undefined || patch.hints !== undefined;
    if (s.rawAnalysis && rematches) rescoreNow(buckets, s.textSettings);
    return { ...s, buckets, analysing: Boolean(s.rawAnalysis) && rematches };
  }), [rescoreNow]);

  const addBucket = useCallback(() => setState((s) => ({
    ...s,
    buckets: [...s.buckets, {
      id: `bucket_${Date.now()}`,
      label: 'New category',
      description: '',
      hints: [],
    }],
  })), []);

  const removeBucket = useCallback((id) => setState((s) => {
    const buckets = s.buckets.filter((b) => b.id !== id);
    if (s.rawAnalysis) rescoreNow(buckets, s.textSettings);
    return { ...s, buckets, analysing: Boolean(s.rawAnalysis) };
  }), [rescoreNow]);

  // Every correction is the same shape: change the overrides record and
  // re-apply it to the raw result. Nothing here touches `rawAnalysis`.
  const editOverrides = useCallback((edit) => setState((s) => {
    if (!s.rawAnalysis) return s;
    const overrides = edit(s.textOverrides ?? EMPTY_OVERRIDES);
    return { ...s, textOverrides: overrides, analysis: applyOverrides(s.rawAnalysis, overrides) };
  }), []);

  const retagFragment = useCallback((fragmentId, bucketId) => editOverrides((o) => ({
    ...o, retags: { ...o.retags, [fragmentId]: bucketId },
  })), [editOverrides]);

  const toggleNoise = useCallback((fragmentId) => editOverrides((o) => ({
    ...o,
    noise: o.noise.includes(fragmentId)
      ? o.noise.filter((id) => id !== fragmentId)
      : [...o.noise, fragmentId],
  })), [editOverrides]);

  const renameTheme = useCallback((themeId, name) => editOverrides((o) => ({
    ...o, themeNames: { ...o.themeNames, [themeId]: name },
  })), [editOverrides]);

  const mergeThemes = useCallback((fromId, intoId) => editOverrides((o) => ({
    ...o, themeMerges: { ...o.themeMerges, [fromId]: intoId },
  })), [editOverrides]);

  const togglePin = useCallback((id) => editOverrides((o) => ({
    ...o, pinned: o.pinned.includes(id) ? o.pinned.filter((x) => x !== id) : [...o.pinned, id],
  })), [editOverrides]);

  const toggleSuppress = useCallback((id) => editOverrides((o) => ({
    ...o,
    suppressed: o.suppressed.includes(id)
      ? o.suppressed.filter((x) => x !== id)
      : [...o.suppressed, id],
  })), [editOverrides]);

  const resetOverrides = useCallback(() => editOverrides(() => EMPTY_OVERRIDES), [editOverrides]);

  const applyAnalysisColumns = useCallback(
    () => setState((s) => ({ ...withAnalysisCharted(s), stage: 'dashboard' })), [],
  );

  const reset = useCallback(() => {
    bytesRef.current = null;
    fileNameRef.current = '';
    setState(IDLE_STATE);
  }, []);

  const setStage = useCallback((stage) => setState((s) => ({ ...s, stage })), []);

  // Any change to the plan or the profile invalidates the dataset, so
  // the apply is re-run from raw rather than patched. That is the whole
  // non-destructive model: nothing ever edits the values in place.
  const { profile, plan, stage, grid } = state;
  useEffect(() => {
    if (stage === 'idle' || stage === 'parsing') return;
    // Carries the grid only when the worker's copy is out of date --
    // when the analysis columns were just appended, say. Every other
    // re-clean stays a plan-sized message.
    const send = gridToSend(grid, workerGridRef.current);
    if (send) workerGridRef.current = send;
    requestClean(profile, plan, send);
  }, [profile, plan, stage, grid, requestClean]);

  /**
   * Read the written answers, unprompted.
   *
   * This is the whole point of the section, so it is not offered behind
   * a button the way it was when it lived inside a general charting
   * tool: a sheet with written answers in it gets them read. Held until
   * the cleaned dataset exists, so the model download starts behind a
   * screen that already has charts on it rather than in front of an
   * empty one. `startAnalysis` latches `autoAnalysed` as it goes and
   * only a new import clears it, which is what stops a re-clean from
   * queueing a second analysis of the same column.
   */
  const { brief, dataset, autoAnalysed, analysing } = state;
  useEffect(() => {
    if (!brief?.analyseColumn || autoAnalysed) return;
    if (!dataset || analysing) return;
    startAnalysis(brief.analyseColumn, { navigate: false });
  }, [brief, dataset, autoAnalysed, analysing, startAnalysis]);

  const value = useMemo(() => ({
    ...state,
    importFile,
    selectSheet,
    setHeaderIndex,
    removeTile,
    cycleTileSize,
    addFilter,
    removeFilter,
    clearFilters,
    selectMark,
    clearSelection,
    startAnalysis,
    showHiddenColumns,
    hideAdminColumns,
    dismissBrief,
    setTextSetting,
    updateBucket,
    addBucket,
    removeBucket,
    retagFragment,
    toggleNoise,
    renameTheme,
    mergeThemes,
    togglePin,
    toggleSuppress,
    resetOverrides,
    applyAnalysisColumns,
    reset,
    setStage,
  }), [
    state, importFile, selectSheet, setHeaderIndex, removeTile, cycleTileSize,
    addFilter, removeFilter, clearFilters, selectMark, clearSelection,
    startAnalysis, showHiddenColumns, hideAdminColumns, dismissBrief,
    setTextSetting, updateBucket, addBucket, removeBucket,
    retagFragment, toggleNoise, renameTheme, mergeThemes, togglePin, toggleSuppress,
    resetOverrides, applyAnalysisColumns,
    reset, setStage,
  ]);

  return <SemanticContext.Provider value={value}>{children}</SemanticContext.Provider>;
}
