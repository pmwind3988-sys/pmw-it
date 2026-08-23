import { StorageFullError } from '../../datastudio/store/db.js';

/**
 * Where a delivery lives between the store room and SharePoint.
 *
 * This is what makes "no signal in the store room" a non-issue rather than a
 * feature: scanning writes here, and only a deliberate Save reaches the
 * network. A batch that has not been saved yet has no half-written state to
 * reason about — it is simply still on the phone.
 *
 * Its own database rather than a store inside Data Studio's, because the two
 * sections have nothing to say to each other and a shared version number would
 * make a schema change in one an upgrade for both.
 */

export const DB_NAME = 'pmw-assets';
export const DB_VERSION = 1;

export const STORE_BATCHES = 'batches';
// Photos are Blobs and are heavy. They live apart from the batch record so
// that listing the unsaved batches does not deserialise a delivery's worth of
// photographs to render a banner saying "2 batches waiting".
export const STORE_PHOTOS = 'photos';

export { StorageFullError };

let dbPromise = null;

export function openDb() {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_BATCHES)) {
        db.createObjectStore(STORE_BATCHES, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_PHOTOS)) {
        db.createObjectStore(STORE_PHOTOS, { keyPath: 'id' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  return dbPromise;
}

/** Tests open a fresh fake database per file; without this they share one. */
export function resetDbForTests() {
  dbPromise = null;
}

/**
 * Every write funnels through here so the quota translation happens in exactly
 * one place. A QuotaExceededError surfaces on the transaction OR on the
 * request depending on the browser, so both are watched.
 */
async function run(storeNames, mode, work) {
  const db = await openDb();

  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeNames, mode);

    const fail = (error) => {
      reject(error?.name === 'QuotaExceededError' ? new StorageFullError() : error);
    };

    tx.oncomplete = () => resolve();
    tx.onerror = () => fail(tx.error);
    tx.onabort = () => fail(tx.error);

    try {
      work(Array.isArray(storeNames)
        ? Object.fromEntries(storeNames.map((name) => [name, tx.objectStore(name)]))
        : tx.objectStore(storeNames));
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

// --- batches ----------------------------------------------------------

export async function saveBatch(batch) {
  const record = { ...batch, updatedAt: Date.now() };
  await run(STORE_BATCHES, 'readwrite', (store) => store.put(record));
  return record;
}

export async function loadBatch(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_BATCHES, 'readonly');
  return promisify(tx.objectStore(STORE_BATCHES).get(id));
}

/** Newest first, which is the order the banner and the list both want. */
export async function listBatches() {
  const db = await openDb();
  const tx = db.transaction(STORE_BATCHES, 'readonly');
  const all = await promisify(tx.objectStore(STORE_BATCHES).getAll());
  return all.sort((a, b) => (b.createdAt ?? 0) - (a.createdAt ?? 0));
}

/**
 * A batch and every photo belonging to it. Deleting the record alone would
 * leave the photos occupying storage that nothing can any longer reach — the
 * one kind of leak a quota-limited store cannot recover from.
 */
export async function deleteBatch(id) {
  const batch = await loadBatch(id);
  const photoIds = photoIdsOf(batch);

  await run([STORE_BATCHES, STORE_PHOTOS], 'readwrite', (stores) => {
    stores[STORE_BATCHES].delete(id);
    for (const photoId of photoIds) stores[STORE_PHOTOS].delete(photoId);
  });
}

export function photoIdsOf(batch) {
  if (!batch) return [];
  const ids = (batch.drafts ?? []).map((draft) => draft.photoId).filter(Boolean);
  if (batch.purchase?.poPhotoId) ids.push(batch.purchase.poPhotoId);
  return ids;
}

// --- photos -----------------------------------------------------------

export async function savePhoto(id, blob) {
  await run(STORE_PHOTOS, 'readwrite', (store) => store.put({ id, blob, savedAt: Date.now() }));
  return id;
}

export async function loadPhoto(id) {
  if (!id) return null;
  const db = await openDb();
  const tx = db.transaction(STORE_PHOTOS, 'readonly');
  const record = await promisify(tx.objectStore(STORE_PHOTOS).get(id));
  return record?.blob ?? null;
}

export async function deletePhoto(id) {
  await run(STORE_PHOTOS, 'readwrite', (store) => store.delete(id));
}

/**
 * How much room the unsaved batches are taking, for the storage-full dialog.
 * A number the user can act on — "delete this delivery and get 40MB back" —
 * rather than a browser error nobody can do anything with.
 */
export async function batchesBySize() {
  const db = await openDb();
  const batches = await listBatches();
  const tx = db.transaction(STORE_PHOTOS, 'readonly');
  const store = tx.objectStore(STORE_PHOTOS);

  const sized = [];
  for (const batch of batches) {
    let bytes = 0;
    for (const photoId of photoIdsOf(batch)) {
      const record = await promisify(store.get(photoId));
      bytes += record?.blob?.size ?? 0;
    }
    sized.push({ ...batch, bytes });
  }

  return sized.sort((a, b) => b.bytes - a.bytes);
}
