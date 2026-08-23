import { useCallback, useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAllAssets } from './sharepoint/readAssets';

export const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The one SharePoint read for this section. The register table, the figures
 * above it and the duplicate check during a review all hang off it, so a count
 * on a card and the rows it opens cannot come from two different fetches.
 */
export function useAssets() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();

  const [assets, setAssets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [loadedAt, setLoadedAt] = useState(null);
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
        const rows = await readAllAssets(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        if (cancelled) return;
        setAssets(rows);
        setLoadedAt(Date.now());
      } catch (failure) {
        if (!cancelled) setError(failure.message || 'Could not load the asset register');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, nonce]);

  return { assets, loading, error, loadedAt, reload };
}
