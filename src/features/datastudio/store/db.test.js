import { describe, it, expect, beforeEach } from 'vitest';
import 'fake-indexeddb/auto';
import {
  saveDataset, loadDataset, listDatasets, deleteDataset,
  saveDashboard, listDashboards, loadDashboard, deleteDashboard,
  saveCleanPlan, loadCleanPlan, datasetsBySize, storageEstimate,
} from './db.js';

const record = () => ({
  id: 'ds1',
  name: 'Requests',
  sourceFileName: 'r.xlsx',
  sheetName: 'Sheet1',
  importedAt: Date.now(),
  rowCount: 3,
  columns: [{ name: 'Amount', type: 'numeric', role: 'measure' }],
  rawColumns: [new Float64Array([1, 2, 3])],
});

async function wipe() {
  for (const d of await listDatasets()) await deleteDataset(d.id);
  for (const d of await listDashboards()) await deleteDashboard(d.id);
}

describe('datasets', () => {
  beforeEach(wipe);

  it('round-trips a dataset', async () => {
    await saveDataset(record());
    expect((await loadDataset('ds1')).name).toBe('Requests');
  });

  it('preserves TypedArrays through structured clone', async () => {
    await saveDataset(record());
    const back = await loadDataset('ds1');
    expect(back.rawColumns[0]).toBeInstanceOf(Float64Array);
    expect(Array.from(back.rawColumns[0])).toEqual([1, 2, 3]);
  });

  it('lists datasets without their column payloads', async () => {
    await saveDataset(record());
    const list = await listDatasets();
    expect(list[0]).toMatchObject({ id: 'ds1', rowCount: 3 });
    expect(list[0].rawColumns).toBeUndefined();
  });

  it('deletes a dataset', async () => {
    await saveDataset(record());
    await deleteDataset('ds1');
    expect(await loadDataset('ds1')).toBeUndefined();
  });

  it('returns undefined for an id that was never saved', async () => {
    expect(await loadDataset('never')).toBeUndefined();
  });

  it('overwrites a dataset saved twice under one id', async () => {
    await saveDataset(record());
    await saveDataset({ ...record(), name: 'Renamed' });
    expect(await listDatasets()).toHaveLength(1);
    expect((await loadDataset('ds1')).name).toBe('Renamed');
  });

  it('lists the most recently imported dataset first', async () => {
    await saveDataset({ ...record(), id: 'old', importedAt: 1000 });
    await saveDataset({ ...record(), id: 'new', importedAt: 2000 });
    expect((await listDatasets()).map((d) => d.id)).toEqual(['new', 'old']);
  });
});

describe('clean plans', () => {
  beforeEach(wipe);

  it('stores plans separately from the dataset', async () => {
    await saveDataset(record());
    await saveCleanPlan('ds1', [{ id: 's1', op: 'trimWhitespace', enabled: true }]);
    expect(await loadCleanPlan('ds1')).toHaveLength(1);
    // The dataset blob itself must be untouched by a plan edit (spec §11).
    expect((await loadDataset('ds1')).rowCount).toBe(3);
  });

  it('returns null when a dataset has no saved plan', async () => {
    expect(await loadCleanPlan('nothing-here')).toBeNull();
  });
});

describe('dashboards', () => {
  beforeEach(wipe);

  it('lists dashboards filtered by dataset', async () => {
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles: [], globalFilters: [] });
    await saveDashboard({ id: 'd2', datasetId: 'ds2', name: 'B', tiles: [], globalFilters: [] });
    expect((await listDashboards('ds1')).map((d) => d.id)).toEqual(['d1']);
  });

  it('round-trips the tiles and filters a dashboard carries', async () => {
    const tiles = [{ id: 't1', chart: 'bar', title: 'X' }];
    const globalFilters = [{ column: 'Dept', kind: 'in', values: ['HR'] }];
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles, globalFilters });
    const back = await loadDashboard('d1');
    expect(back.tiles).toEqual(tiles);
    expect(back.globalFilters).toEqual(globalFilters);
  });

  it('stamps a dashboard with a save time', async () => {
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles: [], globalFilters: [] });
    expect((await loadDashboard('d1')).updatedAt).toBeGreaterThan(0);
  });

  // Orphaned dashboards and plans would occupy quota that nothing can
  // reach, on a screen whose whole job is to let the user free space.
  it('takes a dataset\'s dashboards and plan down with it', async () => {
    await saveDataset(record());
    await saveCleanPlan('ds1', [{ id: 's1', op: 'trimWhitespace', enabled: true }]);
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles: [], globalFilters: [] });

    await deleteDataset('ds1');

    expect(await listDashboards('ds1')).toEqual([]);
    expect(await loadCleanPlan('ds1')).toBeNull();
  });

  it('leaves another dataset\'s dashboards alone', async () => {
    await saveDataset(record());
    await saveDashboard({ id: 'd1', datasetId: 'ds1', name: 'A', tiles: [], globalFilters: [] });
    await saveDashboard({ id: 'd2', datasetId: 'ds2', name: 'B', tiles: [], globalFilters: [] });
    await deleteDataset('ds1');
    expect((await listDashboards()).map((d) => d.id)).toEqual(['d2']);
  });
});

describe('quota reporting', () => {
  beforeEach(wipe);

  it('orders datasets by the bytes their columns occupy', async () => {
    await saveDataset({ ...record(), id: 'small', rawColumns: [new Float64Array(10)] });
    await saveDataset({ ...record(), id: 'big', rawColumns: [new Float64Array(1000)] });
    const sized = await datasetsBySize();
    expect(sized.map((d) => d.id)).toEqual(['big', 'small']);
    expect(sized[0].bytes).toBeGreaterThan(sized[1].bytes);
  });

  // A usage meter must degrade to "unknown", not take the page down,
  // where the Storage API is missing.
  it('reports nulls rather than throwing when the API is unavailable', async () => {
    expect(await storageEstimate()).toMatchObject({ usage: null, quota: null });
  });
});
