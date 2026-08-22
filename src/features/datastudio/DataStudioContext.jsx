import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DataStudioContext, IDLE_STATE } from './dataStudioStore.js';
import { profileColumn } from './profile/profileColumn.js';
import { proposeCleanPlan } from './clean/proposeCleanPlan.js';
import { suggestCharts } from './suggest/suggestCharts.js';
import { SIZE_ORDER } from './dataStudioStore.js';
import {
  saveDataset, loadDataset, saveCleanPlan, loadCleanPlan, StorageFullError,
  saveAnalysis, loadAnalysis,
} from './store/db.js';
import { profileDataset, retopProfile } from './profile/profileDataset.js';
import { detectTextColumns } from './text/detectTextColumns.js';
import { STARTER_BUCKETS } from './text/buckets.js';
import { applyOverrides, EMPTY_OVERRIDES } from './text/overrides.js';
import { deriveColumns, DERIVED_OVERRIDES, DERIVED_HEADERS } from './text/deriveColumns.js';
import { withAnalysisTiles } from './text/analysisTiles.js';
import { gridToSend } from './worker/gridSync.js';
import { planAutopilot } from './intent/planAutopilot.js';
import { hideColumns, unhideColumns } from './intent/hideColumns.js';

/**
 * Owns the worker and the stage machine for the whole section.
 *
 * The worker is created once and terminated on unmount. The bytes of the
 * imported file are kept in a ref, not in state: changing sheet or
 * correcting the header row means re-parsing the same file, and asking
 * the user to drop it again for that would be absurd. They are in a ref
 * rather than state because nothing renders from them and a few MB in a
 * state setter would re-render the tree for no visible reason.
 */
export function DataStudioProvider({ children }) {
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
  // is stale -- the user has ticked something else since -- and applying
  // it would show them the answer to a question they already changed.
  const cleanIdRef = useRef(0);
  // The grid the WORKER is holding. It keeps the parsed grid so a
  // re-clean stays a small message, which means the main thread has to
  // notice when it has replaced that grid -- otherwise the worker
  // cleans the previous sheet and the difference shows up only as tiles
  // reporting a column that is plainly on screen. See `worker/gridSync.js`.
  const workerGridRef = useRef(null);

  useEffect(() => {
    // Vite's native worker syntax -- no plugin, and the worker's own
    // imports (SheetJS included) are bundled into a separate chunk that
    // the main thread never downloads.
    const worker = new Worker(new URL('./worker/studio.worker.js', import.meta.url), {
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
        // Everything the autopilot decides is decided here, once, from
        // the parse result -- not spread across the effects below. The
        // provider's job after this is only to carry the plan out.
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
          // Straight to the canvas. The profile and clean screens are
          // still reachable from the brief, but they are now somewhere
          // the user GOES rather than somewhere they are held: on a
          // survey export the answer to every question those screens ask
          // is the one the autopilot already picked.
          stage: 'canvas',
          brief,
          autoSeed: true,
          autoAnalysed: false,
          sheets: msg.sheets,
          activeSheet: msg.activeSheet,
          headerIndex: msg.headerIndex,
          headerCandidates: msg.headerCandidates ?? [],
          grid: msg.grid,
          profile,
          // A different sheet or header row is a different set of
          // columns, so overrides and the plan from the previous parse
          // no longer refer to anything and are dropped rather than
          // misapplied.
          // The hidden columns go in as ordinary overrides, so the
          // profile panel shows them as overridden and the existing
          // per-column control is all it takes to disagree.
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
          // One shot only: every later clean leaves the canvas alone.
          if (!s.autoSeed) return next;
          return {
            ...next,
            autoSeed: false,
            tiles: s.tiles.length > 0
              ? s.tiles
              : suggestCharts(s.profile, msg.dataset, s.brief?.focus ?? []),
          };
        });
        return;
      }

      if (msg.type === 'error') {
        // Never leave the UI on a spinner (spec §12) -- go back to a
        // screen the user can act on, carrying the reason.
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
   * Constructing it eagerly would pull the model chunk into every Data
   * Studio visit, including the ones that never open the tab.
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
          return {
            ...s,
            rawAnalysis: msg.raw,
            analysis: applyOverrides(msg.raw, overrides),
            textOverrides: overrides,
            analysing: false,
            textError: '',
            textProgress: { stage: '', pct: 100 },
          };
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

  // Re-parsing for a header change also re-detects nothing else: the
  // sheet stays put and only the split between header and body moves.
  const setHeaderIndex = useCallback(
    (index) => setState((s) => {
      post(s.activeSheet, index);
      return s;
    }),
    [post],
  );

  // A type or role override re-profiles ONLY that column. Re-running the
  // whole dataset would discard the other columns' overrides and cost a
  // full pass over every row for a change to one of them.
  const overrideColumn = useCallback((columnName, override) => {
    setState((s) => {
      if (!s.grid || !s.profile) return s;
      const index = s.profile.columns.findIndex((c) => c.name === columnName);
      if (index === -1) return s;

      const values = s.grid.rows.map((row) => row?.[index]);
      const columns = s.profile.columns.slice();
      columns[index] = profileColumn(values, columnName, index, override);
      // Re-topped, or ignoring the column that happened to be the
      // headline measure would leave the profile still naming it.
      const profile = retopProfile({ ...s.profile, columns });

      return {
        ...s,
        overrides: { ...s.overrides, [columnName]: override },
        profile,
        // The plan was proposed against the old verdict, so a column
        // that is no longer numeric must lose its "read as numeric"
        // step rather than keep it and coerce the column back.
        plan: proposeCleanPlan(profile, s.grid),
      };
    });
  }, []);

  /**
   * Put the autopilot's hidden columns back on the canvas.
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

  // --- clean review -----------------------------------------------------

  const setStepEnabled = useCallback((id, enabled) => {
    setState((s) => ({
      ...s,
      plan: s.plan.map((step) => (step.id === id ? { ...step, enabled } : step)),
    }));
  }, []);

  const removeStep = useCallback((id) => {
    setState((s) => ({ ...s, plan: s.plan.filter((step) => step.id !== id) }));
  }, []);

  // A merge the user built by hand. Confidence 'low' because nothing
  // inferred it -- and enabled anyway, because they asked for it.
  const addManualMerge = useCallback((columnName, keys, canonical) => {
    setState((s) => {
      const id = `manualMerge:${columnName}`;
      const existing = s.plan.find((step) => step.id === id);
      const map = { ...(existing?.params?.map ?? {}) };
      for (const key of keys) map[key] = canonical;

      const affectedCount = Object.keys(map).length;
      const merged = {
        id,
        column: columnName,
        op: 'mergeCategories',
        params: { map },
        confidence: 'low',
        affectedCount,
        preview: `Merge ${affectedCount} spellings you picked into "${canonical}"`,
        enabled: true,
        manual: true,
      };

      return {
        ...s,
        plan: existing
          ? s.plan.map((step) => (step.id === id ? merged : step))
          : [...s.plan, merged],
      };
    });
  }, []);

  // The date-order flip for an ambiguous or conflicting column. It has
  // to rewrite the cast step's params rather than add a step, or the
  // column would be cast twice under two different readings.
  const setColumnDateOrder = useCallback((columnName, order) => {
    setState((s) => ({
      ...s,
      dateOrders: { ...s.dateOrders, [columnName]: order },
      plan: s.plan.map((step) => (
        step.column === columnName && step.op === 'castType'
          ? { ...step, params: { ...step.params, order }, enabled: true }
          : step)),
    }));
  }, []);

  // Spec §9 hard rule: date-ONLY columns are never shifted, whatever the
  // toggle says. Adding eight hours to a pure date moves it to the wrong
  // day, which is a data corruption the user cannot see.
  const setColumnZone = useCallback((columnName, sourceZone) => {
    setState((s) => {
      const column = s.profile?.columns.find((c) => c.name === columnName);
      if (column?.type === 'date') return s;
      return {
        ...s,
        zones: { ...s.zones, [columnName]: sourceZone },
        plan: s.plan.map((step) => (
          step.column === columnName && step.op === 'castType'
            ? { ...step, params: { ...step.params, sourceZone } }
            : step)),
      };
    });
  }, []);

  // --- the canvas -------------------------------------------------------

  const commitClean = useCallback(() => setState((s) => ({
    ...s,
    stage: 'canvas',
    // Seed the canvas only when it is empty. Re-seeding would throw away
    // tiles the user has edited every time they stepped back to the
    // clean screen and forward again.
    tiles: s.tiles.length > 0
      ? s.tiles
      : suggestCharts(s.profile, s.dataset, s.brief?.focus ?? []),
  })), []);

  const addTile = useCallback((tile) => setState((s) => ({
    ...s,
    tiles: [...s.tiles, { ...tile, id: tile.id ?? `tile_${Date.now()}_${s.tiles.length}` }],
  })), []);

  const updateTile = useCallback((id, patch) => setState((s) => ({
    ...s,
    tiles: s.tiles.map((t) => (t.id === id ? { ...t, ...patch } : t)),
  })), []);

  const removeTile = useCallback((id) => setState((s) => ({
    ...s,
    tiles: s.tiles.filter((t) => t.id !== id),
    editingTileId: s.editingTileId === id ? null : s.editingTileId,
    // A selection that came from a tile that no longer exists would
    // filter the whole canvas with nothing on screen explaining why.
    selection: s.selection?.sourceTileId === id ? null : s.selection,
  })), []);

  const duplicateTile = useCallback((id) => setState((s) => {
    const index = s.tiles.findIndex((t) => t.id === id);
    if (index === -1) return s;
    const copy = {
      ...s.tiles[index],
      id: `tile_${Date.now()}_${s.tiles.length}`,
      title: `${s.tiles[index].title} (copy)`,
    };
    const tiles = s.tiles.slice();
    tiles.splice(index + 1, 0, copy);
    return { ...s, tiles };
  }), []);

  // Keyboard reordering is the whole reordering story -- there is no
  // drag engine -- so it has to be exact rather than approximate.
  const moveTile = useCallback((id, delta) => setState((s) => {
    const index = s.tiles.findIndex((t) => t.id === id);
    const target = index + delta;
    if (index === -1 || target < 0 || target >= s.tiles.length) return s;
    const tiles = s.tiles.slice();
    [tiles[index], tiles[target]] = [tiles[target], tiles[index]];
    return { ...s, tiles };
  }), []);

  const cycleTileSize = useCallback((id) => setState((s) => ({
    ...s,
    tiles: s.tiles.map((t) => (t.id === id
      ? { ...t, size: SIZE_ORDER[(SIZE_ORDER.indexOf(t.size ?? 'M') + 1) % SIZE_ORDER.length] }
      : t)),
  })), []);

  const setEditingTile = useCallback(
    (id) => setState((s) => ({ ...s, editingTileId: id })), [],
  );

  // --- filters and cross-filter selection --------------------------------

  const addFilter = useCallback((filter) => setState((s) => ({
    ...s,
    // One filter per column: a second filter on the same column would
    // AND with the first and silently produce an empty dashboard.
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
   * A click on a mark. `additive` is a shift-click.
   *
   * Clicking the value that is already the whole selection clears it --
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

  // --- persistence ------------------------------------------------------

  /**
   * Saves the RAW grid, not the cleaned dataset (spec §11).
   *
   * Cleaned columns are derived from raw plus plan on load, so someone
   * who later unticks a cleaning step gets their original values back.
   * Persisting the cleaned blob would bake today's decisions into the
   * file permanently.
   */
  const saveCurrentDataset = useCallback(async (name) => {
    let saved = null;
    setState((s) => { saved = s; return s; });
    // Read the latest state through a no-op setter above, since this
    // runs from an event handler that may hold a stale closure.
    const { grid, profile, plan, fileName, activeSheet, datasetId } = saved ?? {};
    if (!grid || !profile) return null;

    const id = datasetId ?? `ds_${Date.now()}`;
    const rawColumns = profile.columns.map(
      (column) => grid.rows.map((row) => row?.[column.index] ?? null),
    );

    try {
      await saveDataset({
        id,
        name: name || fileName || 'Imported sheet',
        sourceFileName: fileName,
        sheetName: activeSheet,
        importedAt: Date.now(),
        rowCount: grid.rows.length,
        headers: grid.headers,
        columns: profile.columns,
        profile,
        rawColumns,
      });
      await saveCleanPlan(id, plan);
      setState((s) => ({ ...s, datasetId: id, storageFull: false, error: '' }));
      return id;
    } catch (err) {
      setState((s) => ({
        ...s,
        storageFull: err instanceof StorageFullError,
        error: err?.message ?? 'Could not save that dataset.',
      }));
      return null;
    }
  }, []);

  const openSavedDataset = useCallback(async (id) => {
    try {
      const record = await loadDataset(id);
      if (!record) return;

      // Back to a row-major grid, which is what the clean pipeline and
      // the profile panel both read.
      const rows = Array.from({ length: record.rowCount }, (_, r) => (
        record.rawColumns.map((column) => column?.[r] ?? null)));
      const grid = { headers: record.headers, rows };
      const plan = (await loadCleanPlan(id)) ?? proposeCleanPlan(record.profile, grid);

      setState((s) => ({
        ...s,
        stage: 'canvas',
        datasetId: id,
        fileName: record.sourceFileName ?? record.name,
        activeSheet: record.sheetName ?? '',
        sheets: record.sheetName ? [record.sheetName] : [],
        headerCandidates: [],
        grid,
        profile: record.profile,
        plan,
        dataset: null,
        tiles: [],
        globalFilters: [],
        selection: null,
        // A reopened dataset gets starter charts for the same reason a
        // fresh import does -- it was saved from a canvas, and coming
        // back to an empty one reads as data loss.
        autoSeed: true,
        // No brief: nothing was decided this time round. The roles the
        // autopilot set at import are already part of the saved profile.
        brief: null,
        autoAnalysed: false,
        textColumns: detectTextColumns(record.profile, grid),
        error: '',
      }));

      workerGridRef.current = grid;
      requestClean(record.profile, plan, grid);

      // Corrections outlive the session they were made in. The analysis
      // itself is not stored -- it is re-derived from these on demand,
      // which is cheaper than keeping a copy that can drift from the data.
      const saved = await loadAnalysis(id);
      if (saved) {
        setState((s) => ({
          ...s,
          buckets: saved.buckets ?? [],
          textOverrides: saved.overrides ?? EMPTY_OVERRIDES,
          textSettings: saved.settings ?? s.textSettings,
          textColumnName: saved.columnName ?? '',
        }));
      }
    } catch (err) {
      setState((s) => ({ ...s, error: err?.message ?? 'Could not open that dataset.' }));
    }
  }, [requestClean]);

  // Loading a saved dashboard replaces the tiles wholesale. A merge
  // would leave the user with both dashboards at once and no way to
  // tell which tile came from where.
  const applyDashboard = useCallback((record) => setState((s) => ({
    ...s,
    tiles: record?.tiles ?? [],
    globalFilters: record?.globalFilters ?? [],
    selection: null,
    stage: 'canvas',
  })), []);

  // --- text analysis ----------------------------------------------------

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
   * `navigate` is false when the autopilot starts this by itself. The
   * analysis then runs in the background while the user reads their
   * charts, and the canvas shows how far it has got -- being thrown onto
   * a progress bar for something they never asked for would be the
   * feature getting in its own way.
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
    // bucket is presentation, and re-running the model for it would make
    // typing in the name field stutter for no result.
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

  /**
   * Append the analysis to the sheet as five more columns.
   *
   * The grid is re-profiled afterwards rather than patched, so the new
   * columns go through the same type inference as every other column and
   * the canvas needs no special case. `DERIVED_OVERRIDES` supplies the
   * one thing inference cannot know: the categories column is
   * multi-valued by construction, even though most respondents raise a
   * single category and most of its cells therefore carry no separator.
   */
  const applyAnalysisColumns = useCallback(() => setState((s) => {
    if (!s.analysis || !s.grid) return s;

    const { headers, columns } = deriveColumns(s.analysis, s.grid.rows.length);

    // Replace rather than append on a second run, or re-analysing leaves
    // the user with two columns called "Severity".
    const keep = s.grid.headers
      .map((name, i) => ({ name, i }))
      .filter(({ name }) => !DERIVED_HEADERS.includes(name));

    const grid = {
      headers: [...keep.map((k) => k.name), ...headers],
      rows: s.grid.rows.map((row, r) => [
        ...keep.map(({ i }) => row?.[i] ?? null),
        ...columns.map((column) => column[r]),
      ]),
    };
    const profile = profileDataset(grid, DERIVED_OVERRIDES);

    return {
      ...s,
      grid,
      profile,
      plan: proposeCleanPlan(profile, grid),
      textColumns: detectTextColumns(profile, grid),
      // The columns alone landed the user back on the dashboard they
      // already had, with five new columns nothing was charting -- so
      // the button that promised a dashboard of their analysis
      // delivered no visible change at all. The tiles are what makes it
      // a dashboard; the user's own charts are kept below them.
      tiles: withAnalysisTiles(s.tiles, headers),
      stage: 'canvas',
    };
  }), []);

  const dismissStorageFull = useCallback(
    () => setState((s) => ({ ...s, storageFull: false })), [],
  );

  const reset = useCallback(() => {
    bytesRef.current = null;
    fileNameRef.current = '';
    setState(IDLE_STATE);
  }, []);

  const setStage = useCallback((stage) => setState((s) => ({ ...s, stage })), []);

  // Any change to the plan or the profile invalidates the dataset, so
  // the apply is re-run from raw rather than patched. That is the whole
  // non-destructive model: unticking a step must leave no trace of it.
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
   * The autopilot's last step: read the written answers.
   *
   * Held until the cleaned dataset exists, so the model download starts
   * behind a canvas that already has charts on it rather than in front
   * of an empty one. `autoAnalysed` latches immediately and is never
   * cleared except by a new import, which is what stops a re-clean or a
   * tile edit from queueing a second analysis of the same column.
   */
  const { brief, dataset, autoAnalysed, analysing } = state;
  useEffect(() => {
    if (!brief?.autoAnalyse || autoAnalysed) return;
    if (!dataset || analysing) return;
    setState((s) => (s.autoAnalysed ? s : { ...s, autoAnalysed: true }));
    startAnalysis(brief.analyseColumn, { navigate: false });
  }, [brief, dataset, autoAnalysed, analysing, startAnalysis]);

  // Corrections are saved against the dataset, not the file, so they
  // survive a reload -- in this browser only. Nothing here goes to a
  // server. Only meaningful once the dataset has been saved and so has
  // an id to hang them on.
  const {
    datasetId, buckets, textOverrides, textSettings, textColumnName, rawAnalysis,
  } = state;
  useEffect(() => {
    if (!datasetId || !rawAnalysis) return;
    saveAnalysis({
      datasetId,
      columnName: textColumnName,
      buckets,
      overrides: textOverrides ?? EMPTY_OVERRIDES,
      settings: textSettings,
      vectors: rawAnalysis.vectors,
    }).catch(() => {
      // A failed save must not take the tab down. The work is still on
      // screen; only its persistence was lost.
    });
  }, [datasetId, rawAnalysis, buckets, textOverrides, textSettings, textColumnName]);

  const value = useMemo(() => ({
    ...state,
    importFile,
    selectSheet,
    setHeaderIndex,
    overrideColumn,
    setStepEnabled,
    removeStep,
    addManualMerge,
    setColumnDateOrder,
    setColumnZone,
    commitClean,
    addTile,
    updateTile,
    removeTile,
    duplicateTile,
    moveTile,
    cycleTileSize,
    setEditingTile,
    addFilter,
    removeFilter,
    clearFilters,
    selectMark,
    clearSelection,
    saveCurrentDataset,
    openSavedDataset,
    applyDashboard,
    dismissStorageFull,
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
    state, importFile, selectSheet, setHeaderIndex, overrideColumn, setStepEnabled,
    removeStep, addManualMerge, setColumnDateOrder, setColumnZone, commitClean,
    addTile, updateTile, removeTile, duplicateTile, moveTile, cycleTileSize,
    setEditingTile, addFilter, removeFilter, clearFilters, selectMark, clearSelection,
    saveCurrentDataset, openSavedDataset, applyDashboard, dismissStorageFull,
    startAnalysis, showHiddenColumns, hideAdminColumns, dismissBrief,
    setTextSetting, updateBucket, addBucket, removeBucket,
    retagFragment, toggleNoise, renameTheme, mergeThemes, togglePin, toggleSuppress,
    resetOverrides, applyAnalysisColumns,
    reset, setStage,
  ]);

  return <DataStudioContext.Provider value={value}>{children}</DataStudioContext.Provider>;
}
