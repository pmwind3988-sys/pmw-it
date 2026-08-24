/**
 * Disks that belong to IT, not to the machine being scanned.
 *
 * The scan is run from media carried desk to desk — a USB hard disk and a USB
 * flash drive — so those turn up in `Storage Drives` on whichever machines
 * happened to be scanned with them plugged in. Counted as the machine's own
 * storage they do three wrong things at once: they add capacity the machine
 * does not have, they make an all-SSD laptop read "Mixed", and — because
 * Win32_DiskDrive reports them as `Unspecified`, which this project reads as
 * mechanical — they charge the machine 10 risk points for a spinning disk that
 * is not in it.
 *
 * Matching is on the exact model string, not on "WDC" or on the size, so a
 * genuine 1 TB Western Digital disk inside a desktop still counts.
 */
export const IT_TOOL_DRIVES = [
  'WDC WD10 JPVX-60JC3T1',
  // The flash drive the scan is launched from. No internal disk reports this
  // model, so an exact match cannot swallow a machine's real storage.
  'USB DISK 2.0',
];

const normalise = (model) => String(model ?? '').replace(/\s+/g, ' ').trim().toLowerCase();

const KNOWN = new Set(IT_TOOL_DRIVES.map(normalise));

/** True for a disk IT carries — a drive to report but never to count. */
export const isItToolDrive = (model) => KNOWN.has(normalise(model));
