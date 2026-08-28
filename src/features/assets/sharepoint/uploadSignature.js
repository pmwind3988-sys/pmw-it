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

const defaultWait = (ms) => new Promise((resolve) => { setTimeout(resolve, ms); });

/**
 * Uploads one signature and gives back its path, or an empty string when there
 * was nothing to upload. Throwing is left to the caller to catch: what a
 * failure should cost differs between handing out and taking back, and neither
 * of them is "lose the record".
 *
 * Tried more than once, unlike an item photograph. A photograph can be retaken
 * from the thing itself an hour later; a signature exists for the few seconds
 * the person is standing at the desk, and one dropped request on store-room
 * wifi is the difference between a handover somebody signed for and a handover
 * nobody can prove. Three tries with a short wait between them costs a second
 * at worst and saves the signature in the case that actually happens.
 */
export async function uploadSignature({
  siteUrl, token, digest, dataUrl, seed, attempts = 3, wait = defaultWait,
}) {
  const blob = signatureBlob(dataUrl);
  if (!blob) return '';

  let last;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const folder = await photoFolderUrl(siteUrl, token);
      const url = await uploadPhoto(siteUrl, token, digest, {
        folder,
        blob,
        // The moment is part of the name, so a second try never lands on the
        // first try's half-written file and two people signing in the same
        // minute keep two signatures rather than one.
        seed: `signature-${seed || 'handover'}`,
      });

      return url ?? '';
    } catch (thrown) {
      last = thrown;
      if (attempt < attempts) await wait(400 * attempt);
    }
  }

  throw last;
}
