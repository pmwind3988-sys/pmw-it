import { deriveStorage } from './deriveStorage.js';
import { riskScore } from './riskScore.js';
import { enrichFit } from './enrichFit.js';
import { parseReport } from '../parse/parseReport.js';
import { deriveHardwareIds } from './deriveHardwareIds.js';
import { deriveLicensing } from './deriveLicensing.js';
import { deriveSoftware } from './deriveSoftware.js';
import { deriveScanInfo, scanRank } from './deriveScanInfo.js';

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
 * The persona verdict is laid over the row on the same trip, and for the same
 * reason: department, Office products, graphics and mapped drives are all on
 * the record already, so how well the machine suits its desk is worked out
 * fresh rather than stored and left to go stale.
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

/**
 * A version 2 report imported by a build that did not yet know its labels is
 * stored whole, but every new block landed in Extra Fields and the row still
 * reads as an old scan, so its page hides the licences, part numbers and
 * software it plainly holds. The report text is the evidence, so those fields
 * are read back out of it here, the same way the storage summary is.
 *
 * The stored text can also be cut short (a NUL in a program name ended it in
 * SharePoint, before decodeReport removed them), Scan Info and all. Extra
 * Fields is JSON, so it kept each unknown block whole: a block found there
 * wins over the same block in the text, which may be the cut one.
 *
 * Only for a row that says old scan: a row holding a new scan already has what
 * its own import read, and a report that really is v1 has nothing to add.
 */
function extraText(record) {
  let extras;
  try {
    extras = JSON.parse(record.extraFields ?? '[]');
  } catch {
    return '';
  }
  if (!Array.isArray(extras)) return '';
  return extras
    .filter((entry) => entry?.label)
    .map((entry) => `${entry.label}:\n${entry.value ?? ''}\n\n`)
    .join('');
}

function rereadNewScan(record) {
  if (scanRank(record.scanVersion) > 1) return record;
  const raw = String(record.rawReport ?? '');
  const extras = extraText(record);
  if (!/^Scan Info:/m.test(raw + extras)) return record;

  const fromRaw = parseReport(raw);
  const fromExtras = parseReport(extras);
  const fields = { ...fromRaw.fields };
  for (const [label, lines] of Object.entries(fromExtras.fields)) {
    if (lines.length) fields[label] = lines;
  }
  const unknownLabels = [...fromRaw.unknownLabels, ...fromExtras.unknownLabels
    .filter((entry) => !fromRaw.unknownLabels.some((seen) => seen.label === entry.label))];
  const scanInfo = deriveScanInfo(fields);
  if (scanRank(scanInfo.scanVersion) <= 1) return record;

  return {
    ...record,
    ...deriveHardwareIds(fields),
    ...deriveLicensing(fields),
    ...deriveSoftware(fields),
    ...scanInfo,
    extraFields: unknownLabels.length ? JSON.stringify(unknownLabels) : null,
  };
}

export function refixStored(stored) {
  const record = stored ? rereadNewScan(stored) : stored;
  if (!record?.storageDrivesRaw) return enrichFit(record);

  const lines = String(record.storageDrivesRaw)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) return enrichFit(record);

  const storage = deriveStorage(lines);
  if (!storageChanged(record, storage)) return enrichFit(record);

  // The spinning-disk charge and its reason both hang off `hasHdd`, so a drive
  // that just left the count must also leave the risk score — otherwise the
  // page would show an all-SSD machine still charged 10 points for a disk it
  // does not have.
  const next = { ...record, ...storage };
  return enrichFit({ ...next, ...riskScore(next) });
}
