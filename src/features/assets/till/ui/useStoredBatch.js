import { useCallback, useEffect, useState } from 'react';
import { saveBatch, loadBatch, deleteBatch } from '../../store/assetDb';
import { readPref, writePref } from '../prefs';

/**
 * A receipt that is a batch, kept on this device: picked back up after a
 * reload or a dead battery, and listed with the other unsaved deliveries on
 * /assets. `storageKey` remembers WHICH batch is this receipt's, so a delivery
 * and a catch-up count can each be open at once without mixing.
 */
export function useStoredBatch(storageKey, make) {
  const [batch, setBatch] = useState(make);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let alive = true;
    (async () => {
      const id = readPref(storageKey);
      const found = id ? await loadBatch(id).catch(() => null) : null;
      if (!alive) return;
      if (found && found.status !== 'saved') setBatch(found);
      setReady(true);
    })();
    return () => { alive = false; };
  }, [storageKey]);

  useEffect(() => {
    if (!ready || !batch.drafts.length) return;
    writePref(storageKey, batch.id);
    saveBatch(batch).catch(() => {
      // Storage full or blocked: the receipt is still on screen and saving
      // still works; it just will not survive closing the tab.
    });
  }, [batch, ready, storageKey]);

  /** After a full save: the device copy goes, and a fresh receipt starts. */
  const finish = useCallback(async (saved, next) => {
    await deleteBatch(saved.id).catch(() => {});
    writePref(storageKey, null);
    setBatch(next);
  }, [storageKey]);

  return [batch, setBatch, finish];
}
