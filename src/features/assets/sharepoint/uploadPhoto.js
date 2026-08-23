import { spFetch, spUpload, listPath } from '../../sharepoint/spClient.js';
import { PHOTO_LIBRARY_NAME } from './assetSchema.js';

/**
 * Getting a photograph into the document library and getting back a URL the
 * register can point at.
 */

/**
 * The library's real folder path, asked for rather than assumed.
 *
 * A library created with the title "IT Asset Photos" does not reliably live at
 * `/sites/…/IT Asset Photos` — SharePoint derives the folder name from the
 * title in ways that depend on tenant settings and on what was there before.
 * Guessing it produces a 404 on every upload, and the fix is one request.
 */
export async function photoFolderUrl(siteUrl, token) {
  const path = `${listPath(PHOTO_LIBRARY_NAME)}/RootFolder?$select=ServerRelativeUrl`;
  const response = await spFetch(siteUrl, path, { token });
  if (!response.ok) {
    throw new Error(`Could not find the photo library (${response.status})`);
  }

  const data = await response.json();
  const url = data.d?.ServerRelativeUrl;
  if (!url) throw new Error('The photo library reported no folder');
  return url;
}

/**
 * A file name that is stable, readable and legal.
 *
 * SharePoint refuses `" * : < > ? / \ |` and a leading or trailing dot, and it
 * treats `#` and `%` in a path as separators — so anything outside a safe set
 * becomes a dash rather than being left to fail at upload time.
 */
export function photoFileName(seed, { at = Date.now(), extension = 'jpg' } = {}) {
  const slug = String(seed ?? '')
    .replace(/[^A-Za-z0-9._-]+/g, '-')
    .replace(/^[-.]+|[-.]+$/g, '')
    .slice(0, 80)
    .toLowerCase();

  return `${slug || 'item'}-${at}.${extension}`;
}

/**
 * Uploads one photo and returns its server-relative URL.
 *
 * `folder` is the value from `photoFolderUrl`, passed in rather than looked up
 * per photo: a delivery of thirty items would otherwise ask the same question
 * thirty times.
 */
export async function uploadPhoto(siteUrl, token, digest, {
  folder, blob, seed, subfolder = '',
}) {
  if (!blob) return null;

  const target = subfolder ? `${folder}/${subfolder}` : folder;
  const name = photoFileName(seed, { extension: extensionFor(blob.type) });
  const path = `/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(target)}')`
    + `/Files/add(url='${encodeURIComponent(name)}',overwrite=true)`;

  const body = await blob.arrayBuffer();
  const response = await spUpload(siteUrl, path, { token, digest, body });

  if (!response.ok) {
    throw new Error(`Could not upload ${name} (${response.status}): ${await response.text()}`);
  }

  const data = await response.json();
  return data.d?.ServerRelativeUrl ?? `${target}/${name}`;
}

function extensionFor(mimeType) {
  if (mimeType === 'image/png') return 'png';
  if (mimeType === 'image/webp') return 'webp';
  return 'jpg';
}
