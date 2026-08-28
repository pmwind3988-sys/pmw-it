import { photoFolderUrl, uploadPhoto } from './uploadPhoto.js';

/**
 * A signature drawn on a phone, on its way to somewhere everybody can see it.
 *
 * The picture goes into the same library as the item photographs and the
 * handover row keeps the path — a signature is tens of kilobytes of PNG and a
 * SharePoint text column holds 255 characters, so the row can hold where it is
 * and nothing more.
 *
 * Deliberately different from the checklist's signature, which IS the record
 * and takes the whole submission down with it when it fails. Here it is asked
 * for and can be skipped: a laptop that physically changed hands has changed
 * hands whether or not the camera-shy PNG uploaded, and refusing to record
 * that would leave the register lying about where the laptop is.
 */

/** The bytes inside a `data:` URL, or null if that is not what this is. */
export function dataUrlToBytes(dataUrl) {
  const base64 = String(dataUrl ?? '').split(',')[1];
  if (!base64) return null;

  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

/** The same bytes as something `uploadPhoto` will take. */
export function signatureBlob(dataUrl) {
  const bytes = dataUrlToBytes(dataUrl);
  if (!bytes?.length) return null;
  return new Blob([bytes], { type: 'image/png' });
}

/**
 * Uploads one signature and gives back its path, or an empty string when there
 * was nothing to upload. Throwing is left to the caller to catch: what a
 * failure should cost differs between handing out and taking back, and neither
 * of them is "lose the record".
 */
export async function uploadSignature({
  siteUrl, token, digest, dataUrl, seed,
}) {
  const blob = signatureBlob(dataUrl);
  if (!blob) return '';

  const folder = await photoFolderUrl(siteUrl, token);
  const url = await uploadPhoto(siteUrl, token, digest, {
    folder,
    blob,
    seed: `signature-${seed || 'handover'}`,
  });

  return url ?? '';
}
