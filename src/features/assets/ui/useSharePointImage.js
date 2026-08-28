import { useEffect, useState } from 'react';
import { useSharePointToken } from '../../../hooks/useRequests';
import { absoluteFileUrl, fileApiPath } from '../sharepoint/fileUrl';
import { loadImage, peekImage } from './imageCache';

/**
 * A photograph that lives in SharePoint, as something an `<img>` can show.
 *
 * Pointing `src` at the library is not enough, twice over. The browser fetches
 * an image without the app's access token, so SharePoint answers a sign-in
 * page instead of a JPEG; and the library path sends no CORS headers, so
 * asking for it from the portal fails before there is a status code to read.
 * The bytes come through `/_api/`, which takes the token and does send them.
 *
 * States, because a photo failing must not look like a photo missing:
 *   `{ url: null, loading: true }`  — being fetched
 *   `{ url: '<blob:…>' }`           — ready
 *   `{ url: null, failed: true }`   — there is a photo, it could not be shown
 *
 * A picture fetched once is kept for the rest of the visit (`imageCache`), so
 * coming back to an item, or paging back to a unit, shows it at once instead
 * of downloading it again. That is also why nothing is revoked here any more:
 * the cache owns the object URL and revokes it when it falls out.
 */

/** The bytes, fetched with the token SharePoint insists on. */
const fetchBytes = (key, getToken) => async () => {
  const tokenRes = await getToken();
  const response = await fetch(key, {
    headers: { Authorization: `Bearer ${tokenRes.accessToken}` },
  });
  if (!response.ok) throw new Error(String(response.status));
  return response.blob();
};

/** Where the bytes of a stored picture are fetched from, or '' for nothing. */
export function imageKey(siteUrl, stored) {
  const path = fileApiPath(stored);
  return path && siteUrl ? `${siteUrl}${path}` : '';
}

export function useSharePointImage(siteUrl, stored) {
  const getToken = useSharePointToken();
  const href = absoluteFileUrl(siteUrl, stored);
  const key = href ? imageKey(siteUrl, stored) : '';

  const [state, setState] = useState({ key: null, url: null, failed: false });

  useEffect(() => {
    if (!key) return undefined;

    let cancelled = false;

    loadImage(key, fetchBytes(key, getToken))
      .then((url) => { if (!cancelled) setState({ key, url, failed: false }); })
      .catch(() => { if (!cancelled) setState({ key, url: null, failed: true }); });

    return () => { cancelled = true; };
  }, [key, getToken]);

  // Compared during render rather than cleared from an effect: opening a
  // second item must show nothing rather than one frame of the previous item's
  // photograph. A picture already in the cache is picked up here, which is
  // what makes a second look instant rather than another spinner.
  const current = state.key === key ? state : { url: peekImage(key), failed: false };

  return {
    url: current.url,
    href,
    failed: current.failed,
    loading: Boolean(href) && !current.url && !current.failed,
  };
}

/**
 * Start fetching a picture that is not on screen yet — the next unit's
 * photograph while somebody is still reading this one. Free when it is already
 * cached, and a failure here is nobody's business: whatever eventually shows
 * the picture will report it.
 */
export function prefetchSharePointImage(siteUrl, stored, getToken) {
  const key = imageKey(siteUrl, stored);
  if (!key || peekImage(key)) return;

  loadImage(key, fetchBytes(key, getToken)).catch(() => {});
}
