import { useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { isDeviceAdmin } from './sharepoint/deviceAdmin';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * Whether this person may remove machines. False until SharePoint says
 * otherwise, so the remove controls never flash up for somebody who may not
 * use them.
 */
export function useDeviceAdmin() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();
  const [admin, setAdmin] = useState(false);

  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;
    (async () => {
      try {
        const tokenRes = await getToken();
        const yes = await isDeviceAdmin(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        if (!cancelled) setAdmin(yes);
      } catch {
        // Not knowing is not being an admin.
      }
    })();
    return () => { cancelled = true; };
  }, [isAuthenticated, getToken]);

  return admin;
}
