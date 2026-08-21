import { parseDrives } from '../parse/parseValues.js';

export function deriveStorage(driveLines) {
  const drives = parseDrives(driveLines);

  if (!drives.length) {
    return { storageTotalGB: null, driveCount: 0, hasHdd: false, storageType: 'Unknown' };
  }

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
  };
}
