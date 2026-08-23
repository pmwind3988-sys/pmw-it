import { useCallback, useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAllHandovers } from './sharepoint/readHandovers';
import { SHAREPOINT_SITE_URL } from './useAssets';

/**
 * The one handover read for this section. The people list, one person's page,
 * an item's history and the overdue figure all hang off it, so a count on a
 * card and the rows it opens cannot come from two different fetches.
 */
export function useHandovers() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();

  const [handovers, setHandovers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [nonce, setNonce] = useState(0);

  const reload = useCallback(() => setNonce((n) => n + 1), []);

  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;

    (async () => {
      setLoading(true);
      setError('');
      try {
        const tokenRes = await getToken();
        const rows = await readAllHandovers(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        if (!cancelled) setHandovers(rows);
      } catch (failure) {
        if (!cancelled) setError(failure.message || 'Could not load the handovers');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, nonce]);

  return { handovers, loading, error, reload };
}
