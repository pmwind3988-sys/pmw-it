import { useCallback, useEffect, useState } from 'react';
import { listBatches, deleteBatch } from './store/assetDb';
import { BATCH_STATUS } from './draft/batch';

/**
 * The deliveries sitting on this device, unsaved.
 *
 * They are invisible to everybody else until somebody presses Save, which is
 * the price of being able to scan with no signal (§4.1). The banner this feeds
 * is therefore not decoration — it is the only thing standing between a
 * scanned delivery and being forgotten on a phone.
 */
export function useBatches() {
  const [batches, setBatches] = useState([]);
  const [loading, setLoading] = useState(true);
  const [nonce, setNonce] = useState(0);

  const reload = useCallback(() => setNonce((n) => n + 1), []);

  useEffect(() => {
    let cancelled = false;

    (async () => {
      setLoading(true);
      try {
        const all = await listBatches();
        if (!cancelled) setBatches(all.filter((batch) => batch.status !== BATCH_STATUS.SAVED));
      } catch {
        // A browser with IndexedDB blocked (private mode on some platforms) is
        // still a usable register — it just cannot scan offline. Failing the
        // whole page over it would be the wrong trade.
        if (!cancelled) setBatches([]);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [nonce]);

  const discard = useCallback(async (id) => {
    await deleteBatch(id);
    setNonce((n) => n + 1);
  }, []);

  const pendingItems = batches.reduce((sum, batch) => sum + (batch.drafts?.length ?? 0), 0);

  return { batches, pendingItems, loading, reload, discard };
}
