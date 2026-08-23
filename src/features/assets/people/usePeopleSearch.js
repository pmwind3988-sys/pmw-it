import { useEffect, useState } from 'react';
import { useSharePointToken } from '../../../hooks/useRequests';
import { SHAREPOINT_SITE_URL } from '../useAssets';
import { searchPeople, MIN_QUERY } from './peopleSearch';

/** Long enough that typing a name does not fire a request per letter. */
const DEBOUNCE_MS = 300;

/**
 * The directory, searched as somebody types.
 *
 * A failure is returned rather than thrown: the handover screen still accepts a
 * typed name and email when the directory is unreachable, because a search
 * outage must not be the reason a laptop cannot be handed over (§8).
 */
export function usePeopleSearch(query) {
  const getToken = useSharePointToken();
  const [results, setResults] = useState([]);
  const [searching, setSearching] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    const term = String(query ?? '').trim();
    if (term.length < MIN_QUERY) {
      // No request, and nothing to clear: `results` is only ever written by a
      // completed search, and the caller reads it against its own query.
      return undefined;
    }

    let cancelled = false;
    const timer = setTimeout(async () => {
      setSearching(true);
      setError('');
      try {
        const tokenRes = await getToken();
        const people = await searchPeople(SHAREPOINT_SITE_URL, tokenRes.accessToken, term);
        if (!cancelled) setResults(people.map((person) => ({ ...person, forQuery: term })));
      } catch (failure) {
        if (!cancelled) setError(failure.message || 'The directory search failed');
      } finally {
        if (!cancelled) setSearching(false);
      }
    }, DEBOUNCE_MS);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [query, getToken]);

  const term = String(query ?? '').trim();

  return {
    // Compared against the query they were fetched for, so a stale set from a
    // previous term is never shown beside a newer one.
    results: results.filter((person) => person.forQuery === term),
    searching,
    error,
    tooShort: term.length > 0 && term.length < MIN_QUERY,
  };
}
