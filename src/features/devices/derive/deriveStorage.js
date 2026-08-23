import { parseDrives } from '../parse/parseValues.js';
import { isItToolDrive } from './itMedia.js';

const EMPTY = {
  storageTotalGB: null, driveCount: 0, hasHdd: false, storageType: 'Unknown',
};

const describe = (drive) =>
  [drive.model, drive.sizeGB ? `${drive.sizeGB} GB` : null].filter(Boolean).join(' | ');

export function deriveStorage(driveLines) {
  const all = parseDrives(driveLines);

  // The IT extraction disk is reported so the page can say why it is not in
  // the totals, and excluded from every figure derived below.
  const ignored = all.filter((drive) => isItToolDrive(drive.model));
  const drives = all.filter((drive) => !isItToolDrive(drive.model));

  const ignoredDrives = ignored.length ? ignored.map(describe) : null;

  if (!drives.length) return { ...EMPTY, ignoredDrives };

  const sizes = drives.map((drive) => drive.sizeGB).filter((n) => typeof n === 'number');
  const mechanical = drives.filter((drive) => drive.mechanical).length;

  let storageType = 'SSD only';
  if (mechanical === drives.length) storageType = 'HDD only';
  else if (mechanical > 0) storageType = 'Mixed';

  return {
    storageTotalGB: sizes.length ? sizes.reduce((a, b) => a + b, 0) : null,
    driveCount: drives.length,
    hasHdd: mechanical > 0,
    storageType,
    ignoredDrives,
  };
}
