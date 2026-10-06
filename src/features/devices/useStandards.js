import { useCallback, useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readStandards, canEditStandards, pickStandard } from './sharepoint/readStandards';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The standard in force, and whether this person may change it. Until it has
 * loaded -- or if it cannot be read -- the default standard grades the fleet,
 * so a page is never ungraded and never blocked on this read.
 */
export function useStandards() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();
  const [state, setState] = useState({ loaded: pickStandard([]), canEdit: false, loading: true, error: '' });
  const [nonce, setNonce] = useState(0);
  const reload = useCallback(() => setNonce((n) => n + 1), []);

  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;
    (async () => {
      try {
        const tokenRes = await getToken();
        const [loaded, canEdit] = await Promise.all([
          readStandards(SHAREPOINT_SITE_URL, tokenRes.accessToken),
          canEditStandards(SHAREPOINT_SITE_URL, tokenRes.accessToken).catch(() => false),
        ]);
        if (!cancelled) setState({ loaded, canEdit, loading: false, error: '' });
      } catch {
        if (!cancelled) {
          setState((current) => ({
            ...current, loading: false, error: 'Using the default standard: the saved one could not be read.',
          }));
        }
      }
    })();
    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, nonce]);

  return { ...state.loaded, canEdit: state.canEdit, loading: state.loading, error: state.error, reload };
}
