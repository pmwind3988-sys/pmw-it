import { unitsOf } from './units.js';
import { TRACKED } from './assetKinds.js';

/**
 * A delivery entered after the fact, with its paperwork missing.
 *
 * Some of what IT owns arrived before there was anywhere to record it. The
 * delivery note is gone, nobody photographed the boxes, and the serial numbers
 * are on machines already sitting on desks. A register that will not take that
 * delivery until all of it is found is a register that never gets it at all —
 * so the delivery is taken as it stands, MARKED as incomplete, and finished
 * later by whoever walks past the desk.
 *
 * The flag is a delivery-wide switch rather than a field-by-field one: what is
 * missing is the paperwork, and asking about each blank separately turns one
 * decision into thirty.
 */

export const PENDING_YES = 'Yes';
export const PENDING_NO = 'No';

/**
 * A draft carries a boolean; a row read back out of SharePoint carries the
 * word the choice column stores. Every caller would otherwise have to know
 * which of the two it was holding, and the one that forgot would silently
 * report every backfilled row as finished.
 */
export function needsDetails(record) {
  const value = record?.detailsPending;
  if (value === true) return true;
  return String(value ?? '').trim().toLowerCase() === PENDING_YES.toLowerCase();
}

const blank = (value) => !String(value ?? '').trim();

/**
 * What is still to be found, named the way somebody would say it out loud.
 *
 * On a counted line the serials belong to the individual items, so the row is
 * the wrong place to look: asking it would report ten monitors as missing a
 * serial for ever, however many had since been filled in. It counts the empty
 * slots instead, which is the number of machines left to walk up to.
 */
export function missingDetails(record) {
  const missing = [];
  if (blank(record?.doNumber)) missing.push('DO number');

  if (record?.trackingMode === TRACKED) {
    if (blank(record?.serialNumber)) missing.push('Serial number');
  } else {
    const empty = unitsOf(record).filter((unit) => blank(unit.serialNumber)).length;
    if (empty) missing.push(`${empty} serial number${empty === 1 ? '' : 's'}`);
  }

  if (blank(record?.assetTag)) missing.push('Asset label');
  // Either is a picture of the thing: one on this phone, or one already
  // uploaded. A row with neither has never been photographed.
  if (blank(record?.photoUrl) && blank(record?.photoId)) missing.push('Photo');

  return missing;
}
