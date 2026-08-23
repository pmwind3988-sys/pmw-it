import { useEffect, useState } from 'react';
import { loadPhoto } from '../store/assetDb';

/**
 * A stored photo, as something an `<img>` can show.
 *
 * Object URLs are revoked on the way out. Without that, a review grid of
 * thirty photographs leaks thirty blobs every time the list re-renders, and
 * the tab's memory climbs until the phone kills it.
 *
 * The loaded URL is kept together with the id it belongs to, and compared
 * during render. That is what lets switching to a different photo show nothing
 * immediately rather than briefly showing the PREVIOUS item's photograph —
 * without an effect that calls setState just to clear it, which eslint refuses
 * and which costs a second render besides.
 */
export function usePhotoUrl(photoId) {
  const [loaded, setLoaded] = useState({ id: null, url: null });

  useEffect(() => {
    if (!photoId) return undefined;

    let objectUrl = null;
    let cancelled = false;

    loadPhoto(photoId).then((blob) => {
      if (cancelled || !blob) return;
      objectUrl = URL.createObjectURL(blob);
      setLoaded({ id: photoId, url: objectUrl });
    }).catch(() => {});

    return () => {
      cancelled = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [photoId]);

  return loaded.id === photoId ? loaded.url : null;
}
