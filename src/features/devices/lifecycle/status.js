/**
 * Where a machine is in its life. A blank status reads as In use, which is
 * what every row saved before statuses existed actually is -- so the column
 * needed no migration.
 *
 * Only In use and In repair count towards the fleet's figures. A retired
 * laptop reported as a "critical risk" would be a figure lying about the
 * machines people are working on.
 */
export const IN_USE = 'In use';
export const IN_REPAIR = 'In repair';
export const SPARE = 'Spare';
export const RETIRED = 'Retired';

export const STATUSES = [IN_USE, IN_REPAIR, SPARE, RETIRED];

export function statusOf(device) {
  return STATUSES.includes(device?.status) ? device.status : IN_USE;
}

export function inFleet(device) {
  const status = statusOf(device);
  return status === IN_USE || status === IN_REPAIR;
}
