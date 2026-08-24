import { getFormDigest } from '../../sharepoint/spClient.js';
import { photoFolderUrl, uploadPhoto } from './uploadPhoto.js';
import { parseUnits, serialiseUnits } from '../units.js';

/**
 * The photographs taken of individual items, on their way to SharePoint.
 *
 * A photo taken in the pager is written to the phone first and carried on the
 * unit record as `photoId` — a local reference, not a URL. That is what makes
 * photographing a shelf in a store room with no signal work at all, and it is
 * also why something has to swap those references for real library paths
 * before the row is written. This is that something.
 *
 * A photo that will not upload does NOT take the save down with it. The
 * serial number somebody just typed is the record; the picture is an
 * attachment to it, and losing the first because the second produced an odd
 * JPEG would be the wrong trade. The reference is left in place so the next
 * save tries again, and the failure is reported back.
 */
export async function uploadUnitPhotos({
  siteUrl, token, stored, seed, photoFor,
}) {
  const units = parseUnits(stored);
  const pending = units.filter((unit) => unit.photoId);
  // The overwhelmingly common case: nothing new was photographed, so not one
  // request is spent finding a folder to put nothing in.
  if (!pending.length) return { units: stored, uploaded: 0, failures: [] };

  const digest = await getFormDigest(siteUrl, token);
  const folder = await photoFolderUrl(siteUrl, token);

  const failures = [];
  let uploaded = 0;

  const next = [];
  for (const unit of units) {
    if (!unit.photoId) {
      next.push(unit);
      continue;
    }

    const blob = await photoFor?.(unit.photoId);
    if (!blob) {
      // The reference points at nothing — the phone's storage was cleared, or
      // the photo was taken on another device. Dropped rather than retried
      // for ever, because no future save can make it appear.
      next.push({ ...unit, photoId: '' });
      continue;
    }

    try {
      const url = await uploadPhoto(siteUrl, token, digest, {
        folder,
        blob,
        seed: `${seed || 'item'}-unit-${unit.index + 1}`,
      });
      next.push({ ...unit, photoUrl: url, photoId: '' });
      uploaded += 1;
    } catch (error) {
      failures.push({ index: unit.index, error: error.message });
      next.push(unit);
    }
  }

  return { units: serialiseUnits(next), uploaded, failures };
}
