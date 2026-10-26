import { hardwareIdsOf } from './derive/deriveHardwareIds.js';
import { softwareOf } from './derive/deriveSoftware.js';

/**
 * Plain facts read off a machine for the licence, serial and software views:
 * the columns that show them, the filters that narrow by them and the finder
 * that searches them. Pure, so each rule is tested without a table.
 */

/** Which scan wrote the report. An old report (or one with no version) is unknown, never wrong. */
export function scanKindOf(version) {
  if (version === null || version === undefined || version === '') return 'old';
  const text = String(version);
  if (/basic$/i.test(text)) return 'basic';
  const number = Number.parseInt(text, 10);
  if (!Number.isFinite(number) || number <= 1) return 'old';
  return 'new';
}

export const SCAN_LABELS = { new: 'New scan', basic: 'Basic scan', old: 'Old scan' };
export const SCAN_BADGES = { new: 'New', basic: 'Basic', old: 'Old' };

export const isOldScan = (device) => scanKindOf(device?.scanVersion) === 'old';

/**
 * "Personal copy", "Personal account" or both, from the reasons the derive
 * step wrote. Null when the machine is not flagged. A flag with no reason
 * text is read as a personal copy, the commoner of the two.
 */
export function officeWarningOf(device) {
  if (device?.officePersonal !== true) return null;
  const reason = String(device.officePersonalReason ?? '');
  const copy = !reason || /personal (Microsoft 365|Office)/i.test(reason);
  const account = /personal Microsoft account/i.test(reason);
  if (copy && account) return 'Personal copy and account';
  if (account) return 'Personal account';
  return 'Personal copy';
}

/** "2 × M471A5244CB0-CWE", one entry per part number, so a pair reads as one line. */
export function memoryPartsOf(device) {
  const counts = new Map();
  for (const module of hardwareIdsOf(device).memoryModules) {
    const part = module.partNumber || module.maker || 'Unknown module';
    counts.set(part, (counts.get(part) ?? 0) + 1);
  }
  return [...counts].map(([part, count]) => (count > 1 ? `${count} × ${part}` : part));
}

/** One serial per drive; a drive whose scan gave no serial says so rather than vanishing. */
export function driveSerialsOf(device) {
  return hardwareIdsOf(device).driveSerials
    .map((drive) => drive.serial ?? (drive.model ? `${drive.model} (no serial)` : null))
    .filter(Boolean);
}

/** Installed programs whose name contains the text, case-insensitively. */
export function programsMatching(device, text) {
  const needle = String(text ?? '').trim().toLowerCase();
  if (!needle) return [];
  return softwareOf(device).filter((program) => program.name && program.name.toLowerCase().includes(needle));
}

export const hasSoftware = (device, text) => programsMatching(device, text).length > 0;
