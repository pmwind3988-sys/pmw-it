import { scanRank } from './derive/deriveScanInfo.js';
import { scanKindOf } from './fleetFacts.js';

/**
 * The import review's scan figures: how many reports are new, basic or old,
 * and which of them are OLDER than the report already stored for the machine.
 * An older report keeps the stored licences and software (carryV2Fields in
 * sharepoint/syncDevices.js applies the same rule when saving).
 */

/** True when the incoming scan ranks below the stored one, so its licences and software are kept. */
export function isOlderThanStored(incomingVersion, storedVersion) {
  return scanRank(incomingVersion) < scanRank(storedVersion);
}

/** `matches` is matchIncoming's output: { device, existing }. */
export function scanSummary(matches) {
  const summary = { reports: matches.length, new: 0, basic: 0, old: 0, olderThanStored: 0 };
  for (const { device, existing } of matches) {
    summary[scanKindOf(device.scanVersion)] += 1;
    if (existing && isOlderThanStored(device.scanVersion, existing.scanVersion)) {
      summary.olderThanStored += 1;
    }
  }
  return summary;
}

/** "3 Oct" in Malaysia time, or null when there is no instant. */
function shortDay(epochMs) {
  const time = Number(epochMs);
  if (epochMs === null || epochMs === undefined || !Number.isFinite(time)) return null;
  return new Intl.DateTimeFormat('en-GB', {
    day: 'numeric', month: 'short', timeZone: 'Asia/Kuala_Lumpur',
  }).format(new Date(time));
}

/** The one line above the review grid. */
export function scanSummaryLine(summary) {
  const reports = `${summary.reports} ${summary.reports === 1 ? 'report' : 'reports'} read`;
  return `${reports} · ${summary.new} new scan · ${summary.basic} basic scan · ${summary.old} old scan`
    + ` · ${summary.olderThanStored} older than what is stored — licences and software kept`;
}

/** The note under a row whose incoming scan is older than the stored one, or null. */
export function olderScanNote(match) {
  const { device, existing } = match;
  if (!existing || !isOlderThanStored(device.scanVersion, existing.scanVersion)) return null;
  const when = shortDay(existing.scannedOn);
  return `Stored scan is newer (v${existing.scanVersion}${when ? `, ${when}` : ''}). Licences and software kept.`;
}
