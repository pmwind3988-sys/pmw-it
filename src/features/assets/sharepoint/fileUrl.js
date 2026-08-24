/**
 * Turning what the register stored into something a browser can open.
 *
 * `uploadPhoto` records the SERVER-RELATIVE path SharePoint answers with —
 * `/sites/IThelpdesk/IT Asset Photos/laptop-1724.jpg`. Handed straight to an
 * `<img src>` that resolves against the app's OWN origin, so the browser asks
 * the portal for a file only SharePoint has and shows a broken image. Every
 * photograph in the register was invisible for that one missing prefix.
 *
 * A path also has to survive being a URL. Library titles contain spaces, and a
 * raw space in `src` is quietly mangled by some engines and rejected by others.
 */

/** Percent-encode each path segment, leaving the separators alone. */
function encodePath(path) {
  return path
    .split('/')
    .map((segment) => encodeURIComponent(decodeURIComponent(segment)))
    .join('/');
}

/**
 * `stored` is whatever is in `PhotoUrl` or `PoPhotoUrl`; `siteUrl` is the site
 * the register lives on. Answers with an absolute URL, or an empty string when
 * there is nothing to show — never with a half-formed one, because a broken
 * `<img>` is indistinguishable from a photo nobody took.
 */
export function absoluteFileUrl(siteUrl, stored) {
  const value = String(stored ?? '').trim();
  if (!value) return '';

  // Already absolute: an older row, or a link somebody pasted in by hand.
  if (/^https?:\/\//i.test(value)) return value;

  let origin;
  try {
    origin = new URL(siteUrl).origin;
  } catch {
    return '';
  }

  // Server-relative, which is what an upload answers with.
  if (value.startsWith('/')) return `${origin}${encodePath(value)}`;

  // Anything else is relative to the site itself.
  const base = String(siteUrl).replace(/\/+$/, '');
  return `${base}/${encodePath(value)}`;
}

/**
 * The path the app can actually FETCH the bytes from.
 *
 * Not the same URL as the one above, and this is the crux of it: the library
 * path serves the file to a browser carrying SharePoint cookies, and answers a
 * cross-origin fetch from the portal with nothing at all — no CORS headers, so
 * the request fails before a status code exists. `/_api/` does send them, and
 * takes the app's access token, which is the only credential the portal has.
 *
 * So the picture on screen comes through the API and the "open it in
 * SharePoint" link goes to the library, and neither can be swapped for the
 * other.
 */
export function fileApiPath(stored) {
  const value = String(stored ?? '').trim();
  if (!value) return '';

  let serverRelative = value;
  if (/^https?:\/\//i.test(value)) {
    try {
      serverRelative = new URL(value).pathname;
    } catch {
      return '';
    }
  }
  if (!serverRelative.startsWith('/')) return '';

  // The apostrophe is encoded by hand: `encodeURIComponent` leaves it alone,
  // being legal in a URL, and one in a file name would close the OData string
  // literal it sits inside and take the rest of the request with it.
  const encoded = encodePath(serverRelative).replace(/'/g, '%27');

  return `/_api/web/GetFileByServerRelativeUrl('${encoded}')/$value`;
}
