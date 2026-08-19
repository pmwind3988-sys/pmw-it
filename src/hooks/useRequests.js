import { useCallback, useEffect, useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { fetchAllListItems, fetchAllColumnChoices } from '../services/sharePointService';
import { sharePointRequest } from '../authConfig';

export const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

export const CHOICE_COLUMNS = [
  'Entity',
  'Equipment_x0020_Items',
  'Software_x0020_Licenses',
  'Request_x0020_Type',
  'Department',
];

/**
 * One SharePoint token, acquired the same way everywhere: silently, falling
 * back to a popup only when Azure AD says interaction is actually required.
 */
export function useSharePointToken() {
  const { instance } = useMsal();

  return useCallback(async () => {
    const account = instance.getActiveAccount();
    if (!account) throw new Error('No signed-in account');
    try {
      return await instance.acquireTokenSilent({ ...sharePointRequest, account });
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        return await instance.acquireTokenPopup({ ...sharePointRequest, account });
      }
      throw e;
    }
  }, [instance]);
}

/**
 * The request list and the column choices behind it — the one read both the
 * dashboard and the records screen are built on, so the figures on the cards
 * and the rows they drill into can never come from two different fetches.
 *
 * `reload` re-runs it; `loadedAt` is what the dashboard reports as the snapshot
 * time.
 */
export function useRequests() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();

  const [items, setItems] = useState([]);
  const [choices, setChoices] = useState({});
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
        const [itemsData, choiceMap] = await Promise.all([
          fetchAllListItems(SHAREPOINT_SITE_URL, tokenRes.accessToken),
          fetchAllColumnChoices(SHAREPOINT_SITE_URL, tokenRes.accessToken, CHOICE_COLUMNS),
        ]);
        if (cancelled) return;
        setItems(itemsData);
        setChoices(choiceMap);
        setLoadedAt(new Date());
      } catch (err) {
        if (!cancelled) setError(err.message || 'Failed to load data');
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [isAuthenticated, getToken, nonce]);

  return { items, choices, loading, error, loadedAt, reload };
}

/* --- Shared shape of a request row -------------------------------------- */

/** The list's date column, whose static name is unreadable everywhere else. */
export const DATE_FIELD = 'Join_x0020__x002f__x0020_Last_x0';

/**
 * Multi-choice columns come back as a bare array or as `{ results: [...] }`
 * depending on which OData flavour the call asked for, and both shapes reach
 * these screens.
 */
export function toChoiceArray(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.results)) return value.results;
  return [];
}

export function requestDate(item) {
  const raw = item?.[DATE_FIELD] || item?.Created;
  if (!raw) return null;
  const d = new Date(raw);
  return Number.isNaN(d.getTime()) ? null : d;
}

export function formatDate(value) {
  const d = value instanceof Date ? value : requestDate({ [DATE_FIELD]: value });
  return d ? d.toLocaleDateString() : '-';
}
