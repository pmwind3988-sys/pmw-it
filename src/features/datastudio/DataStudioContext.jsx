import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { DataStudioContext, IDLE_STATE } from './dataStudioStore.js';
import { profileColumn } from './profile/profileColumn.js';

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
          // columns, so type overrides from the previous parse no longer
          // refer to anything and are dropped rather than misapplied.
          overrides: {},
          error: '',
          progress: { stage: '', pct: 100 },
        }));
        return;
      }
      if (msg.type === 'error') {
        // Never leave the UI on a spinner (spec §12) -- go back to a
        // screen the user can act on, carrying the reason.
        setState((s) => ({ ...s, stage: 'idle', error: msg.message, progress: { stage: '', pct: 0 } }));
      }
    };

    worker.onerror = (event) => {
      setState((s) => ({
        ...s,
        stage: 'idle',
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
      const nextOverrides = { ...s.overrides, [columnName]: override };
      const columns = s.profile.columns.slice();
      columns[index] = profileColumn(values, columnName, index, override);

      return { ...s, overrides: nextOverrides, profile: { ...s.profile, columns } };
    });
  }, []);

  const reset = useCallback(() => {
    bytesRef.current = null;
    setState(IDLE_STATE);
  }, []);

  const setStage = useCallback((stage) => setState((s) => ({ ...s, stage })), []);

  const value = useMemo(() => ({
    ...state, importFile, selectSheet, setHeaderIndex, overrideColumn, reset, setStage,
  }), [state, importFile, selectSheet, setHeaderIndex, overrideColumn, reset, setStage]);

  return <DataStudioContext.Provider value={value}>{children}</DataStudioContext.Provider>;
}
