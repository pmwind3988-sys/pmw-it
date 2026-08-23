const VERBOSE = 'application/json;odata=verbose';
const NOMETADATA = 'application/json;odata=nometadata';

export const ITEM_ACCEPT = NOMETADATA;
export const listPath = (name) => `/_api/web/lists/getByTitle('${encodeURIComponent(name)}')`;

export function spFetch(siteUrl, path, {
  token, digest, method = 'GET', body, accept = VERBOSE, headers: extra,
}) {
  const headers = {
    Accept: accept,
    'Content-Type': accept,
    Authorization: `Bearer ${token}`,
    ...extra,
  };
  // Applied after the spread so a caller passing its own headers cannot drop
  // the digest and turn every write into a 403.
  if (digest) headers['X-RequestDigest'] = digest;

  return fetch(`${siteUrl}${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
}

/**
 * A file, not an item. `spFetch` JSON-stringifies whatever body it is given,
 * which turns a photograph into the string "[object Blob]" — so a binary upload
 * needs its own door rather than an extra flag on that one.
 *
 * `Content-Type` is deliberately absent: SharePoint reads the bytes, and a
 * declared type here has been observed to make it reject the upload.
 */
export function spUpload(siteUrl, path, { token, digest, body }) {
  return fetch(`${siteUrl}${path}`, {
    method: 'POST',
    headers: {
      Accept: VERBOSE,
      Authorization: `Bearer ${token}`,
      'X-RequestDigest': digest,
    },
    body,
  });
}

export async function getFormDigest(siteUrl, token) {
  const response = await spFetch(siteUrl, '/_api/contextinfo', { token, method: 'POST' });
  if (!response.ok) throw new Error(`Could not get a form digest (${response.status})`);

  const data = await response.json();
  const digest = data?.d?.GetContextWebInformation?.FormDigestValue;
  if (!digest) throw new Error('SharePoint returned no form digest');
  return digest;
}
