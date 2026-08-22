// IndexedDB persistence -- spec §11.
//
// TypedArrays survive structured clone natively, so a column persists
// as-is with no serialisation step and no precision loss. That is the
// whole reason the columnar store is worth having on disk as well as in
// memory.
//
// Only RAW columns are stored. Cleaned columns are derived from raw plus
// plan on load, so a user who changes their mind about a cleaning step
// is not stuck with a blob that was cleaned the old way.

export const DB_NAME = 'pmw-datastudio';
export const DB_VERSION = 2;

export const STORE_DATASETS = 'datasets';
// The heavy column blobs live apart from the dataset metadata.
//
// This splits what spec §11 describes as one store, and it is
// deliberate: IndexedDB has no partial read, so listing datasets from a
// single store would deserialise every Float64Array of every dataset to
// render a list of names. The plan requires that the list not do that.
// `saveDataset` splits the record and `loadDataset` rejoins it, so
// callers never see the seam.
export const STORE_COLUMNS = 'datasetColumns';
export const STORE_PLANS = 'cleanPlans';
export const STORE_DASHBOARDS = 'dashboards';
// Text analysis: the bucket definitions in force, the user's corrections,
// and the settings they were produced under. Never the analysis itself --
// that is derived, and re-deriving it is cheap compared with storing a
// copy that can silently disagree with the data it came from.
export const STORE_ANALYSES = 'analyses';

/**
 * Thrown when the browser refuses a write for want of space.
 *
 * A distinct type because the UI has to react differently: quota is the
 * one storage failure the user can actually fix, and spec §11 requires a
 * "here are your datasets by size" dialog rather than the silent failure
 * that is the default behaviour.
 */
export class StorageFullError extends Error {
  constructor(message = 'There is no room left in this browser\'s storage.') {
    super(message);
    this.name = 'StorageFullError';
  }
}

let dbPromise = null;

export function openDb() {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_DATASETS)) {
        db.createObjectStore(STORE_DATASETS, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_COLUMNS)) {
        db.createObjectStore(STORE_COLUMNS, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_PLANS)) {
        db.createObjectStore(STORE_PLANS, { keyPath: 'datasetId' });
      }
      if (!db.objectStoreNames.contains(STORE_ANALYSES)) {
        db.createObjectStore(STORE_ANALYSES, { keyPath: 'datasetId' });
      }
      if (!db.objectStoreNames.contains(STORE_DASHBOARDS)) {
        const dashboards = db.createObjectStore(STORE_DASHBOARDS, { keyPath: 'id' });
        // So "the dashboards for this dataset" is an index lookup rather
        // than a scan of every dashboard the user has ever saved.
        dashboards.createIndex('datasetId', 'datasetId', { unique: false });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  return dbPromise;
}

// Every write funnels through here so the quota translation happens in
// exactly one place. A QuotaExceededError surfaces on the transaction OR
// the request depending on the browser, so both are watched.
async function run(storeNames, mode, work) {
  const db = await openDb();

  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeNames, mode);
    let result;

    const fail = (error) => {
      reject(error?.name === 'QuotaExceededError' ? new StorageFullError() : error);
    };

    tx.oncomplete = () => resolve(result);
    tx.onerror = () => fail(tx.error);
    tx.onabort = () => fail(tx.error);

    try {
      result = work(
        Array.isArray(storeNames)
          ? Object.fromEntries(storeNames.map((n) => [n, tx.objectStore(n)]))
          : tx.objectStore(storeNames),
        (value) => { result = value; },
      );
    } catch (error) {
      fail(error);
    }
  });
}

function promisify(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

// --- datasets ---------------------------------------------------------

export async function saveDataset(record) {
  const { rawColumns, ...meta } = record;
  await run([STORE_DATASETS, STORE_COLUMNS], 'readwrite', (stores) => {
    stores[STORE_DATASETS].put(meta);
    stores[STORE_COLUMNS].put({ id: record.id, rawColumns: rawColumns ?? [] });
  });
  return record.id;
}

export async function loadDataset(id) {
  const db = await openDb();
  const tx = db.transaction([STORE_DATASETS, STORE_COLUMNS], 'readonly');
  const meta = await promisify(tx.objectStore(STORE_DATASETS).get(id));
  if (!meta) return undefined;
  const columns = await promisify(tx.objectStore(STORE_COLUMNS).get(id));
  return { ...meta, rawColumns: columns?.rawColumns ?? [] };
}

// Metadata only. The column blobs are in their own store precisely so
// this never touches them.
export async function listDatasets() {
  const db = await openDb();
  const tx = db.transaction(STORE_DATASETS, 'readonly');
  const all = await promisify(tx.objectStore(STORE_DATASETS).getAll());
  return all.sort((a, b) => (b.importedAt ?? 0) - (a.importedAt ?? 0));
}

export async function deleteDataset(id) {
  // Everything that belongs to the dataset goes with it. Leaving the
  // clean plan or the dashboards behind would occupy quota that nothing
  // can ever reach or delete.
  const db = await openDb();
  const dashboardIds = (await listDashboards(id)).map((d) => d.id);

  await new Promise((resolve, reject) => {
    const tx = db.transaction(
      [STORE_DATASETS, STORE_COLUMNS, STORE_PLANS, STORE_DASHBOARDS, STORE_ANALYSES],
      'readwrite',
    );
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
    tx.objectStore(STORE_DATASETS).delete(id);
    tx.objectStore(STORE_COLUMNS).delete(id);
    tx.objectStore(STORE_PLANS).delete(id);
    tx.objectStore(STORE_ANALYSES).delete(id);
    const dashboards = tx.objectStore(STORE_DASHBOARDS);
    for (const dashboardId of dashboardIds) dashboards.delete(dashboardId);
  });
}

// --- clean plans ------------------------------------------------------

export async function saveCleanPlan(datasetId, steps) {
  await run(STORE_PLANS, 'readwrite', (store) => {
    store.put({ datasetId, steps, updatedAt: Date.now() });
  });
}

export async function loadCleanPlan(datasetId) {
  const db = await openDb();
  const tx = db.transaction(STORE_PLANS, 'readonly');
  const record = await promisify(tx.objectStore(STORE_PLANS).get(datasetId));
  return record?.steps ?? null;
}

// --- dashboards -------------------------------------------------------

export async function saveDashboard(record) {
  const now = Date.now();
  const value = { createdAt: now, ...record, updatedAt: now };
  await run(STORE_DASHBOARDS, 'readwrite', (store) => {
    store.put(value);
  });
  return value.id;
}

export async function loadDashboard(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_DASHBOARDS, 'readonly');
  return promisify(tx.objectStore(STORE_DASHBOARDS).get(id));
}

export async function listDashboards(datasetId) {
  const db = await openDb();
  const tx = db.transaction(STORE_DASHBOARDS, 'readonly');
  const store = tx.objectStore(STORE_DASHBOARDS);
  const all = datasetId
    ? await promisify(store.index('datasetId').getAll(datasetId))
    : await promisify(store.getAll());
  return all.sort((a, b) => (b.updatedAt ?? 0) - (a.updatedAt ?? 0));
}

export async function deleteDashboard(id) {
  await run(STORE_DASHBOARDS, 'readwrite', (store) => {
    store.delete(id);
  });
}

// --- text analysis ----------------------------------------------------

// Above this, the fragment vectors are re-computed on reopen instead of
// stored. 2,000 x 384 floats is about 3MB, which is a reasonable thing
// to keep; ten times that is not, and a survey that large is rare enough
// that paying for the model run again is the better trade.
export const MAX_CACHED_VECTORS = 2000;

export async function saveAnalysis(record) {
  const vectors = record?.vectors ?? null;
  const value = {
    datasetId: record.datasetId,
    columnName: record.columnName ?? '',
    buckets: record.buckets ?? [],
    overrides: record.overrides ?? {},
    settings: record.settings ?? {},
    vectors: vectors && vectors.length <= MAX_CACHED_VECTORS ? vectors : null,
    updatedAt: Date.now(),
  };
  await run(STORE_ANALYSES, 'readwrite', (store) => {
    store.put(value);
  });
}

export async function loadAnalysis(datasetId) {
  const db = await openDb();
  const tx = db.transaction(STORE_ANALYSES, 'readonly');
  return promisify(tx.objectStore(STORE_ANALYSES).get(datasetId));
}

// --- quota ------------------------------------------------------------

/**
 * How much of the origin's storage allowance is in use.
 *
 * Returns nulls rather than throwing where the API is missing, so a
 * usage meter degrades to "unknown" instead of taking the page down.
 */
export async function storageEstimate() {
  if (typeof navigator === 'undefined' || !navigator.storage?.estimate) {
    return { usage: null, quota: null, ratio: null };
  }
  try {
    const { usage, quota } = await navigator.storage.estimate();
    return {
      usage: usage ?? null,
      quota: quota ?? null,
      ratio: usage && quota ? usage / quota : null,
    };
  } catch {
    return { usage: null, quota: null, ratio: null };
  }
}

// Datasets ordered by how much room they take, which is what a
// "storage full" dialog has to show for its delete buttons to be a
// decision rather than a guess.
export async function datasetsBySize() {
  const db = await openDb();
  const metas = await listDatasets();
  const tx = db.transaction(STORE_COLUMNS, 'readonly');
  const store = tx.objectStore(STORE_COLUMNS);

  const sized = [];
  for (const meta of metas) {
    const columns = await promisify(store.get(meta.id));
    const bytes = (columns?.rawColumns ?? []).reduce(
      (total, column) => total + (column?.byteLength ?? 0), 0,
    );
    sized.push({ ...meta, bytes });
  }

  return sized.sort((a, b) => b.bytes - a.bytes);
}
