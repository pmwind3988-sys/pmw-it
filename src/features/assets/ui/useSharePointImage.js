import { useEffect, useState } from 'react';
import { useSharePointToken } from '../../../hooks/useRequests';
import { absoluteFileUrl, fileApiPath } from '../sharepoint/fileUrl';

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
 * Object URLs are revoked on the way out; without that, paging through units
 * leaks a blob per photograph until the tab is killed.
 */
export function useSharePointImage(siteUrl, stored) {
  const getToken = useSharePointToken();
  const href = absoluteFileUrl(siteUrl, stored);

  const [state, setState] = useState({ href: null, url: null, failed: false });

  useEffect(() => {
    if (!href) return undefined;

    let objectUrl = null;
    let cancelled = false;

    (async () => {
      try {
        const tokenRes = await getToken();
        const response = await fetch(`${siteUrl}${fileApiPath(stored)}`, {
          headers: { Authorization: `Bearer ${tokenRes.accessToken}` },
        });
        if (!response.ok) throw new Error(String(response.status));

        const blob = await response.blob();
        if (cancelled) return;

        objectUrl = URL.createObjectURL(blob);
        setState({ href, url: objectUrl, failed: false });
      } catch {
        if (!cancelled) setState({ href, url: null, failed: true });
      }
    })();

    return () => {
      cancelled = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [href, siteUrl, stored, getToken]);

  // Compared during render rather than cleared from an effect, the same as
  // `usePhotoUrl`: opening a second item must show nothing rather than one
  // frame of the previous item's photograph.
  const current = state.href === href ? state : { url: null, failed: false };

  return {
    url: current.url,
    href,
    failed: current.failed,
    loading: Boolean(href) && !current.url && !current.failed,
  };
}
