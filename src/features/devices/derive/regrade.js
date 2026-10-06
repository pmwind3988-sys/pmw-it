import { partGrades } from './partGrades.js';

/** The fleet graded against one standard. Pure: the rows are not changed. */
export function regrade(devices, standard) {
  return devices.map((device) => ({ ...device, ...partGrades(device, standard) }));
}
