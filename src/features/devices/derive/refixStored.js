import { deriveStorage } from './deriveStorage.js';
import { riskScore } from './riskScore.js';

/**
 * A stored row, brought back in line with today's rules.
 *
 * The storage summary on a row — its type, its drive count, whether it has a
 * spinning disk — is only ever as right as the exclusion rules were the day it
 * was imported. A machine scanned before its IT extraction disk was known
 * carries "Mixed / Has HDD / 2 drives" forever, even though the very same page
 * already greys that disk out of the drive list, because the summary is read
 * straight from SharePoint and the greying is done fresh on the screen.
 *
 * The raw drive list is the evidence and it is stored beside the summary. So
 * rather than wait for a re-scan, the summary is recomputed from those raw
 * lines on the way out of SharePoint, applying today's `isItToolDrive` rules to
 * yesterday's rows. The row self-corrects the moment it is opened.
 *
 * Nothing is written back: the correction lives on the record the page reads,
 * and the stored value is left for the next real sync to overwrite. Re-reading
 * is cheap and safe; a write on every read is neither.
 */

const STORAGE_KEYS = ['storageTotalGB', 'driveCount', 'hasHdd', 'storageType'];

/** True when the recomputed storage disagrees with what was stored. */
function storageChanged(record, storage) {
  return STORAGE_KEYS.some((key) => record[key] !== storage[key]);
}

export function refixStored(record) {
  if (!record?.storageDrivesRaw) return record;

  const lines = String(record.storageDrivesRaw)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) return record;

  const storage = deriveStorage(lines);
  if (!storageChanged(record, storage)) return record;

  // The spinning-disk charge and its reason both hang off `hasHdd`, so a drive
  // that just left the count must also leave the risk score — otherwise the
  // page would show an all-SSD machine still charged 10 points for a disk it
  // does not have.
  const next = { ...record, ...storage };
  return { ...next, ...riskScore(next) };
}
