import { useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from './useRequests';
import { loadOrgDirectory } from '../features/forms/sharepoint/loadOrgDirectory';
import { companyOptions } from '../features/forms/orgDirectory';

/**
 * HR's company and department lists, for a signed-in page.
 *
 * `null` until they arrive; `false` if they could not be read — the form then
 * falls back to the built-in entities and a typed department rather than
 * refusing to open. Shared by the checklist and the link builder, so both
 * offer the same choices.
 */
export function useOrgDirectory() {
  const getToken = useSharePointToken();
  const isAuthenticated = useIsAuthenticated();
  const [directory, setDirectory] = useState(null);

  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;

    (async () => {
      try {
        const tokenRes = await getToken();
        const loaded = await loadOrgDirectory(tokenRes.accessToken);
        if (!cancelled) setDirectory(companyOptions(loaded.companies).length ? loaded : false);
      } catch {
        if (!cancelled) setDirectory(false);
      }
    })();

    return () => { cancelled = true; };
  }, [isAuthenticated, getToken]);

  return directory;
}
