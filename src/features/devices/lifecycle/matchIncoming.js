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
  // Two reports in one import would both update the same machine.
  DUPLICATE_TARGET: 'duplicateTarget',
};

/** Kinds that save nothing: the review says why, and the save skips them. */
export const SKIPPED_KINDS = new Set([MATCH.DUPLICATE_SERIAL, MATCH.DUPLICATE_TARGET]);

/** A choice in `targets` (keyed by source file name): a new machine, or a row id to update. */
export const NEW_MACHINE = 'new';

const nameKey = (name) => String(name ?? '').trim().toLowerCase();
const serialKey = (serial) => (serial ? normaliseSerial(serial) : null);

export function matchIncoming(incoming, existing, targets = {}) {
  return oneClaimEach(withChoices(autoMatch(incoming, existing), existing, targets));
}

function autoMatch(incoming, existing) {
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

/**
 * What IT chose in the review outranks what the serial and the name say. The
 * automatic match is kept as `auto`, so the review can still offer it back.
 */
function withChoices(matches, existing, targets) {
  return matches.map((match) => {
    const auto = match.existing ?? null;
    const choice = targets[match.device.sourceFileName];

    if (choice === NEW_MACHINE) {
      let note = 'Saved as a new machine, as chosen.';
      if (auto) {
        note += ` ${auto.computerName} is left as it is.`;
        const serial = serialKey(match.device.serialNumber);
        if (serial && serial === serialKey(auto.serialNumber)) {
          note += ` Both carry serial ${auto.serialNumber}, so a later report with it will match only one of them.`;
        }
      }
      return { ...match, existing: null, kind: MATCH.NEW, chosen: true, auto, note };
    }

    const row = choice === undefined || choice === null ? null : existing.find((item) => item.id === choice);
    if (row) {
      if (auto?.id === row.id && match.kind !== MATCH.DUPLICATE_SERIAL) return { ...match, chosen: true, auto };
      const renamed = nameKey(row.computerName) !== nameKey(match.device.computerName);
      return {
        ...match,
        existing: row,
        kind: renamed ? MATCH.RENAMED : MATCH.SAME,
        chosen: true,
        auto,
        note: `Updates ${row.computerName}, as chosen${renamed ? ` -- it takes the name ${match.device.computerName}` : ''}.`,
      };
    }

    // A choice of a machine that has since been removed falls back to the automatic match.
    return { ...match, auto };
  });
}

/**
 * One report per machine. A choice is served first, then the automatic
 * matches in file order; a later claim on a machine already taken saves
 * nothing, or two reports would both write the same row and one would be lost.
 */
function oneClaimEach(matches) {
  const claimedBy = new Map();
  const order = [...matches.keys()].sort((a, b) => Number(Boolean(matches[b].chosen)) - Number(Boolean(matches[a].chosen)));
  const result = [...matches];
  for (const index of order) {
    const match = matches[index];
    const id = match.existing?.id;
    if (id === null || id === undefined) continue;
    const first = claimedBy.get(id);
    if (first === undefined) {
      claimedBy.set(id, match.device.computerName ?? match.device.sourceFileName);
      continue;
    }
    result[index] = {
      ...match,
      existing: null,
      kind: MATCH.DUPLICATE_TARGET,
      note: `${match.existing.computerName} is already updated by ${first} in this import, so this report is not saved. Choose another machine or New machine.`,
    };
  }
  return result;
}

/** The line the review shows under a row, or null when there is nothing to say. */
export function noticeFor(match) {
  if (match.note) return match.note;
  if (!match.existing) return null;
  const status = statusOf(match.existing);
  if (status !== SPARE && status !== RETIRED) return null;
  const movedOn = match.existing.statusChangedOn;
  const scannedOn = match.device.scannedOn;
  if (movedOn !== null && movedOn !== undefined && scannedOn !== null && scannedOn !== undefined && scannedOn <= movedOn) {
    return `Older than its move to ${status === SPARE ? 'IT Stash' : 'the Graveyard'}: specs updated, status kept.`;
  }
  if (status === SPARE) return 'Brought back from IT Stash: it goes back into use with whoever the scan names.';
  if (status === RETIRED) return 'Brought back from the Graveyard: it goes back into use with whoever the scan names.';
  return null;
}
