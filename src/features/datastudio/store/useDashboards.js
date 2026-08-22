import { useCallback, useEffect, useState } from 'react';
import {
  listDatasets, deleteDataset, listDashboards, saveDashboard, deleteDashboard,
  storageEstimate, datasetsBySize,
} from './db.js';

// Reloads are driven by bumping a token rather than by calling an async
// function from the effect body. That keeps the first setState behind an
// await -- React warns about synchronous setState in an effect, and the
// warning is right: it costs a second render pass every mount.
function useReloadToken() {
  const [token, setToken] = useState(0);
  const reload = useCallback(() => setToken((n) => n + 1), []);
  return [token, reload];
}

/**
 * The saved-dataset library, plus the storage meter that sits under it.
 *
 * Every call is wrapped so a failing IndexedDB -- private browsing, a
 * corrupt database, a browser with storage switched off -- surfaces as a
 * message on screen rather than an unhandled rejection. This section
 * still works with no persistence at all: the library is simply empty
 * and everything else carries on.
 */
export function useDatasetLibrary() {
  const [datasets, setDatasets] = useState([]);
  const [estimate, setEstimate] = useState({ usage: null, quota: null, ratio: null });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [token, refresh] = useReloadToken();

  useEffect(() => {
    // Guards against a resolved read landing after the component has
    // gone, which would set state on nothing.
    let cancelled = false;

    (async () => {
      try {
        const [saved, usage] = await Promise.all([listDatasets(), storageEstimate()]);
        if (cancelled) return;
        setDatasets(saved);
        setEstimate(usage);
        setError('');
      } catch (err) {
        if (!cancelled) setError(err?.message ?? 'Could not read saved datasets.');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [token]);

  const remove = useCallback(async (id) => {
    try {
      await deleteDataset(id);
      refresh();
    } catch (err) {
      setError(err?.message ?? 'Could not delete that dataset.');
    }
  }, [refresh]);

  const bySize = useCallback(async () => {
    try {
      return await datasetsBySize();
    } catch {
      return [];
    }
  }, []);

  return { datasets, estimate, loading, error, refresh, remove, bySize };
}

/** Saved dashboards for one dataset. */
export function useDashboards(datasetId) {
  const [dashboards, setDashboards] = useState([]);
  const [error, setError] = useState('');
  const [token, refresh] = useReloadToken();

  useEffect(() => {
    let cancelled = false;

    (async () => {
      if (!datasetId) {
        if (!cancelled) setDashboards([]);
        return;
      }
      try {
        const saved = await listDashboards(datasetId);
        if (cancelled) return;
        setDashboards(saved);
        setError('');
      } catch (err) {
        if (!cancelled) setError(err?.message ?? 'Could not read saved dashboards.');
      }
    })();

    return () => { cancelled = true; };
  }, [datasetId, token]);

  const save = useCallback(async (name, tiles, globalFilters) => {
    if (!datasetId) return null;
    try {
      const id = `dash_${Date.now()}`;
      await saveDashboard({ id, datasetId, name, tiles, globalFilters });
      refresh();
      return id;
    } catch (err) {
      setError(err?.message ?? 'Could not save that dashboard.');
      return null;
    }
  }, [datasetId, refresh]);

  const remove = useCallback(async (id) => {
    try {
      await deleteDashboard(id);
      refresh();
    } catch (err) {
      setError(err?.message ?? 'Could not delete that dashboard.');
    }
  }, [refresh]);

  return { dashboards, error, save, remove, refresh };
}
