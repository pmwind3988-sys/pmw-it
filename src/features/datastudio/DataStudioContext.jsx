import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DataStudioContext, IDLE_STATE } from './dataStudioStore.js';
import { profileColumn } from './profile/profileColumn.js';
import { proposeCleanPlan } from './clean/proposeCleanPlan.js';
import { suggestCharts } from './suggest/suggestCharts.js';
import { SIZE_ORDER } from './dataStudioStore.js';

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
  const bytesRef = useRef(null);
  // Monotonic id per clean request. A result whose id is not the latest
  // is stale -- the user has ticked something else since -- and applying
  // it would show them the answer to a question they already changed.
  const cleanIdRef = useRef(0);

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
        setState((s) => ({
          ...s,
          stage: 'profiled',
          sheets: msg.sheets,
          activeSheet: msg.activeSheet,
          headerIndex: msg.headerIndex,
          headerCandidates: msg.headerCandidates ?? [],
          grid: msg.grid,
          profile: msg.profile,
          // A different sheet or header row is a different set of
          // columns, so overrides and the plan from the previous parse
          // no longer refer to anything and are dropped rather than
          // misapplied.
          overrides: {},
          plan: proposeCleanPlan(msg.profile, msg.grid),
          dataset: null,
          error: '',
          progress: { stage: '', pct: 100 },
        }));
        return;
      }

      if (msg.type === 'cleaned') {
        setState((s) => (
          msg.requestId === cleanIdRef.current
            ? { ...s, dataset: msg.dataset, cleaning: false }
            : s));
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
    };
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
  const requestClean = useCallback((profile, plan) => {
    if (!workerRef.current || !profile || !plan) return;
    cleanIdRef.current += 1;
    setState((s) => ({ ...s, cleaning: true }));
    workerRef.current.postMessage({
      type: 'clean', profile, plan, requestId: cleanIdRef.current,
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
      const profile = { ...s.profile, columns };

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
    tiles: s.tiles.length > 0 ? s.tiles : suggestCharts(s.profile, s.dataset),
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

  const reset = useCallback(() => {
    bytesRef.current = null;
    setState(IDLE_STATE);
  }, []);

  const setStage = useCallback((stage) => setState((s) => ({ ...s, stage })), []);

  // Any change to the plan or the profile invalidates the dataset, so
  // the apply is re-run from raw rather than patched. That is the whole
  // non-destructive model: unticking a step must leave no trace of it.
  const { profile, plan, stage } = state;
  useEffect(() => {
    if (stage === 'idle' || stage === 'parsing') return;
    requestClean(profile, plan);
  }, [profile, plan, stage, requestClean]);

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
    reset,
    setStage,
  }), [
    state, importFile, selectSheet, setHeaderIndex, overrideColumn, setStepEnabled,
    removeStep, addManualMerge, setColumnDateOrder, setColumnZone, commitClean,
    addTile, updateTile, removeTile, duplicateTile, moveTile, cycleTileSize,
    setEditingTile, addFilter, removeFilter, clearFilters, selectMark, clearSelection,
    reset, setStage,
  ]);

  return <DataStudioContext.Provider value={value}>{children}</DataStudioContext.Provider>;
}
