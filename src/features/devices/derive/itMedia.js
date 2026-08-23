/**
 * Disks that belong to IT, not to the machine being scanned.
 *
 * The scan is run from a USB disk carried desk to desk, so that disk turns up
 * in `Storage Drives` on whichever machines happened to be scanned with it
 * plugged in. Counted as the machine's own storage it does three wrong things
 * at once: it adds ~932 GB the machine does not have, it makes an all-SSD
 * laptop read "Mixed", and — because Win32_DiskDrive reports it as
 * `Unspecified`, which this project reads as mechanical — it charges the
 * machine 10 risk points for a spinning disk that is not in it.
 *
 * Matching is on the exact model string, not on "WDC" or on the size, so a
 * genuine 1 TB Western Digital disk inside a desktop still counts.
 */
export const IT_TOOL_DRIVES = ['WDC WD10 JPVX-60JC3T1'];

const normalise = (model) => String(model ?? '').replace(/\s+/g, ' ').trim().toLowerCase();

const KNOWN = new Set(IT_TOOL_DRIVES.map(normalise));

/** True for the IT extraction disk — a drive to report but never to count. */
export const isItToolDrive = (model) => KNOWN.has(normalise(model));
