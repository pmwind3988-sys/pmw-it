import { useCallback, useEffect, useState } from 'react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAssignments } from './sharepoint/readAssignments';
import { readChanges } from './sharepoint/readHistory';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/** Who had this machine and what changed on it. Keyed on the id, so a rename does not re-fetch. */
export function useDeviceHistory(device) {
  const getToken = useSharePointToken();
  const [state, setState] = useState({ stints: [], changes: [], loading: true, error: '' });
  const [nonce, setNonce] = useState(0);
  const reload = useCallback(() => setNonce((n) => n + 1), []);
  const id = device?.id ?? null;
  const name = device?.computerName ?? '';

  useEffect(() => {
    if (id === null) return undefined;
    let cancelled = false;
    (async () => {
      try {
        const tokenRes = await getToken();
        const [stints, changes] = await Promise.all([
          readAssignments(SHAREPOINT_SITE_URL, tokenRes.accessToken, { deviceId: id }),
          readChanges(SHAREPOINT_SITE_URL, tokenRes.accessToken, { id, computerName: name }),
        ]);
        if (!cancelled) setState({ stints, changes, loading: false, error: '' });
      } catch (failure) {
        if (!cancelled) setState((current) => ({ ...current, loading: false, error: failure.message }));
      }
    })();
    return () => { cancelled = true; };
  }, [id, name, getToken, nonce]);

  return { ...state, reload };
}
