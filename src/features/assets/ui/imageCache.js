/**
 * Pictures already fetched, kept for the rest of the visit.
 *
 * A photograph in this app is not a `src` the browser can cache for us: the
 * bytes come through `/_api/` carrying an access token, so every `<img>` is
 * really a `fetch` this app makes by hand — and a hand-made fetch gets no help
 * from the browser's own image cache. Without something here, opening an item,
 * going back and opening it again downloads the same 300KB again, and paging
 * through ten units of a delivery re-downloads the delivery-order scan ten
 * times. That is the whole of why photographs felt slow.
 *
 * So the object URL is kept against the path it came from, and the second ask
 * is answered immediately with no request at all. Two pictures asked for at
 * once — the same delivery photo on three thumbnails — share one request
 * rather than racing three.
 *
 * The cache is bounded and the oldest entries are revoked as they fall out.
 * Never revoking would leak a blob per photograph for as long as the tab is
 * open, which is the leak the per-component cleanup used to prevent; keeping
 * the newest sixty is what makes a picture instant without that cost.
 */

const LIMIT = 60;

/** path → object URL, in least-recently-used order (Map preserves insertion). */
const ready = new Map();
/** path → the in-flight promise, so the same picture is fetched once. */
const pending = new Map();

/** Reading counts as using: what is being looked at should not be evicted. */
function touch(key) {
  const url = ready.get(key);
  if (url === undefined) return null;
  ready.delete(key);
  ready.set(key, url);
  return url;
}

function keep(key, url) {
  ready.set(key, url);

  while (ready.size > LIMIT) {
    const [oldest, oldestUrl] = ready.entries().next().value;
    ready.delete(oldest);
    URL.revokeObjectURL(oldestUrl);
  }
}

/** What is already here, or null. Never starts a request. */
export function peekImage(key) {
  if (!key) return null;
  return touch(key);
}

/**
 * The picture at `key`, fetched with `load` only if it is not already here.
 * `load` answers with a Blob; anything else, including a failed request, is
 * left to throw so the caller can say "there is a photo and it did not load"
 * rather than "there is no photo".
 */
export function loadImage(key, load) {
  const cached = touch(key);
  if (cached) return Promise.resolve(cached);

  const already = pending.get(key);
  if (already) return already;

  const request = load()
    .then((blob) => {
      const url = URL.createObjectURL(blob);
      keep(key, url);
      return url;
    })
    .finally(() => pending.delete(key));

  pending.set(key, request);
  return request;
}

/** Emptied between tests; nothing in the app needs to forget a photograph. */
export function clearImageCache() {
  for (const url of ready.values()) URL.revokeObjectURL(url);
  ready.clear();
  pending.clear();
}
