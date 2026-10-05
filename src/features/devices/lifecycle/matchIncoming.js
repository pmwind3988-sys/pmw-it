import { normaliseSerial } from '../parse/placeholders.js';
import { statusOf, SPARE, RETIRED } from './status.js';

/**
 * Which register row an incoming report is about. Serial first, because the
 * computer name is the one thing IT changes when a machine changes hands;
 * name second, because every report scanned before 2026-10-05 has no serial.
 */
export const MATCH = {
  NEW: 'new',
  SAME: 'same',
  RENAMED: 'renamed',
  NAME_REUSED: 'nameReused',
  DUPLICATE_SERIAL: 'duplicateSerial',
};

const nameKey = (name) => String(name ?? '').trim().toLowerCase();
const serialKey = (serial) => (serial ? normaliseSerial(serial) : null);

export function matchIncoming(incoming, existing) {
  const bySerial = new Map();
  const byName = new Map();
  for (const row of existing) {
    const serial = serialKey(row.serialNumber);
    if (serial) bySerial.set(serial, row);
    if (row.computerName) byName.set(nameKey(row.computerName), row);
  }

  const claimed = new Map();

  return incoming.map((device) => {
    const serial = serialKey(device.serialNumber);

    if (serial) {
      const first = claimed.get(serial);
      if (first) {
        return {
          device, existing: null, kind: MATCH.DUPLICATE_SERIAL,
          note: `Same serial as ${first} in this import, so only that one is saved.`,
        };
      }
      claimed.set(serial, device.computerName ?? device.sourceFileName);

      const hit = bySerial.get(serial);
      if (hit) {
        const renamed = nameKey(hit.computerName) !== nameKey(device.computerName);
        return {
          device, existing: hit, kind: renamed ? MATCH.RENAMED : MATCH.SAME,
          note: renamed ? `Renamed from ${hit.computerName}: same serial, so its history carries over.` : null,
        };
      }
    }

    const named = byName.get(nameKey(device.computerName));
    if (named) {
      const theirs = serialKey(named.serialNumber);
      if (serial && theirs && theirs !== serial) {
        return {
          device, existing: null, kind: MATCH.NAME_REUSED,
          note: `${named.computerName} was a different machine (serial ${named.serialNumber}). This one is saved as a new machine.`,
        };
      }
      return { device, existing: named, kind: MATCH.SAME, note: null };
    }

    return { device, existing: null, kind: MATCH.NEW, note: null };
  });
}

/** The line the review shows under a row, or null when there is nothing to say. */
export function noticeFor(match) {
  if (match.note) return match.note;
  if (!match.existing) return null;
  const status = statusOf(match.existing);
  if (status === SPARE) return 'Brought back from IT Stash: it goes back into use with whoever the scan names.';
  if (status === RETIRED) return 'Brought back from the Graveyard: it goes back into use with whoever the scan names.';
  return null;
}
