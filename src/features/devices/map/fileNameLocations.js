import { locationsIn } from './locations.js';
import { KNOWN_DEPARTMENTS } from '../derive/deriveIdentity.js';

/**
 * The locations a file NAME may be read for. A hand-typed "E" or "ADMIN" must
 * not start claiming pieces of file names, so one-character codes and anything
 * that is also a department are left out.
 */
// A file of its own: locations.js is imported by the derive layer, which
// deriveIdentity.js belongs to, so importing it there would be a cycle.
export function fileNameLocations(devices = []) {
  const departments = new Set(KNOWN_DEPARTMENTS);
  return locationsIn(devices).filter((code) => code.length >= 2 && !departments.has(code));
}
