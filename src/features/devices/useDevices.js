import { useCallback, useEffect, useMemo, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAllDevices } from './sharepoint/readDevices';
import { useStandards } from './useStandards';
import { regrade } from './derive/regrade';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The one SharePoint read for this section. The register and the dashboard
 * both hang off it, so a figure on a card and the rows it opens cannot come
 * from two different fetches.
 */
export function useDevices() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();

  const [devices, setDevices] = useState([]);
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
        const rows = await readAllDevices(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        if (!cancelled) setDevices(rows);
      } catch (failure) {
        if (!cancelled) setError(failure.message || 'Could not load the device list');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, nonce]);

  const standards = useStandards();
  // Graded on the way out with the standard in force: saving a standard
  // regrades every machine on the next render, with no re-scan.
  const graded = useMemo(() => regrade(devices, standards.standard), [devices, standards.standard]);

  return { devices: graded, loading, error, reload, standards };
}
