/**
 * Reading a label until it stops changing, and putting the result into a
 * form without destroying what is already in it.
 *
 * Both halves are pure and live here rather than in the hook, for the
 * same reason `scanSession.js` does: no camera is under test, but every
 * decision taken about what comes out of one is.
 *
 * ## Why a value has to be read twice
 *
 * A barcode carries a checksum — a decoder either reads it or does not.
 * Text has no such thing. A hand-held camera on a printed label will
 * read `8` as `B`, `0` as `O` and `1` as `l` in one frame out of
 * several, and no part of the answer says which frame that was. Two
 * frames in a row agreeing is not proof, but it removes the single
 * misread that a one-shot capture would write into the register.
 */

/** The fields a photographed label can fill. */
export const SCAN_FIELDS = [
  'serialNumber', 'partNumber', 'macAddress', 'assetTag',
  'manufacturer', 'model', 'specSummary',
];

/** How many passes must agree before a value is accepted. */
export const AGREE = 2;

/**
 * When to stop. A label held at an angle, or one printed too small for
 * the camera, produces a different answer every pass and would otherwise
 * keep the camera running until the battery went. Eight passes is a few
 * seconds of holding still; after that the honest thing is to say so and
 * offer the keyboard.
 */
export const MAX_PASSES = 8;

export function newTextScan() {
  return {
    passes: 0,
    /** `{ field: { value, count } }` — what the last passes have been saying. */
    pending: {},
    /** `{ field: value }` — what has been read the same way often enough. */
    settled: {},
    guessed: [],
    additional: [],
    /**
     * `{ field, value }` pairs somebody has crossed out. Remembered because
     * the camera is still pointed at the same label: without this, a refusal
     * would be undone by the very next pass, half a second later.
     *
     * Keyed on the PAIR, not the field. Crossing one out means "not that one",
     * not "stop reading this field" -- the usual reason it is wrong is that
     * the camera misread it, and the right value is the next thing to arrive.
     */
    rejected: [],
    exhausted: false,
  };
}

const isRejected = (scan, field, value) => (scan.rejected ?? []).some(
  (entry) => entry.field === field && entry.value === value,
);

/** `fields` is one pass of `readTextFields`. */
export function recordReading(scan, fields) {
  const pending = { ...scan.pending };
  const settled = { ...scan.settled };
  const guessed = new Set(scan.guessed);
  const additional = [...scan.additional];

  for (const field of SCAN_FIELDS) {
    const value = String(fields?.[field] ?? '').trim();
    if (!value) continue;

    // A field that has already settled is left alone. Re-opening it would
    // let the camera drifting onto the next box in the pile overwrite the
    // answer that was already good.
    if (settled[field]) continue;
    if (isRejected(scan, field, value)) continue;

    const previous = pending[field];
    const count = previous?.value === value ? previous.count + 1 : 1;
    pending[field] = { value, count };

    if (count >= AGREE) {
      settled[field] = value;
      delete pending[field];
      if (fields.guessed?.includes(field)) guessed.add(field);
    }
  }

  for (const value of fields?.additional ?? []) {
    if (!value || additional.includes(value)) continue;
    if (isRejected(scan, ADDITIONAL, value)) continue;
    additional.push(value);
  }

  const passes = scan.passes + 1;

  return {
    passes,
    pending,
    settled,
    guessed: [...guessed],
    additional,
    rejected: scan.rejected ?? [],
    exhausted: passes >= MAX_PASSES && Object.keys(pending).length > 0,
  };
}

/** The name the rejected list files a loose line of writing under. */
const ADDITIONAL = '_additional';

/**
 * What the camera is offering, for a person to accept or cross out.
 *
 * Only what has SETTLED. A value still being read has been seen once and may
 * yet turn out to be a misread of something else, and offering it would put
 * the reader back in the business of judging half-finished guesses.
 */
export function candidates(scan) {
  return SCAN_FIELDS
    .filter((field) => scan.settled[field])
    .map((field) => ({
      field,
      value: scan.settled[field],
      guessed: (scan.guessed ?? []).includes(field),
    }));
}

/** Crossed out: off the list, and not offered again. */
export function rejectValue(scan, field) {
  const value = scan.settled[field] ?? scan.pending[field]?.value;
  if (!value) return scan;

  const settled = { ...scan.settled };
  const pending = { ...scan.pending };
  delete settled[field];
  delete pending[field];

  return {
    ...scan,
    settled,
    pending,
    guessed: (scan.guessed ?? []).filter((name) => name !== field),
    rejected: [...(scan.rejected ?? []), { field, value }],
  };
}

/** The same, for a line of writing it read but could not name. */
export function dismissExtra(scan, value) {
  return {
    ...scan,
    additional: (scan.additional ?? []).filter((entry) => entry !== value),
    rejected: [...(scan.rejected ?? []), { field: ADDITIONAL, value }],
  };
}

export function isSettled(scan, field) {
  return Boolean(scan.settled[field]);
}

export function settledValues(scan) {
  return { ...scan.settled };
}

/** Nothing more is going to change: every line being read has been accepted. */
export function isComplete(scan) {
  return scan.passes > 0
    && Object.keys(scan.settled).length > 0
    && Object.keys(scan.pending).length === 0;
}

/**
 * A value the person filled in themselves outranks anything a camera
 * reads — the same contract `setDraftField` and the device import both
 * keep. A value an earlier SCAN guessed does not: correcting a bad guess
 * by looking again is exactly what re-scanning is for.
 */
function canFill(record, field) {
  if (record.manualFields?.includes(field)) return false;
  if (!String(record[field] ?? '').trim()) return true;
  return Boolean(record.guessed?.includes(field));
}

/**
 * Returns the updated record, and the values it refused to write over so
 * the screen can offer them rather than swallow them.
 */
export function applyScannedFields(
  record, values, guessedFields = [], additional = [], { byHand = false } = {},
) {
  const next = { ...record };
  const guessed = new Set(record.guessed ?? []);
  const manual = new Set(record.manualFields ?? []);
  const heldBack = [];

  for (const field of SCAN_FIELDS) {
    const value = String(values?.[field] ?? '').trim();
    if (!value) continue;

    if (!canFill(record, field)) {
      // Only worth mentioning if it differs; re-reading the same value
      // off the same box is not a clash anybody needs to be told about.
      if (String(record[field] ?? '').trim() !== value) heldBack.push({ field, value });
      continue;
    }

    next[field] = value;
    // Ticked off a list rather than written in by a scan that decided for
    // itself. A deliberate choice outranks the next scan exactly as typing it
    // would, or the camera drifting onto the next box in the pile undoes the
    // decision somebody just made on purpose -- and it is no longer a guess,
    // because a person looked at it.
    if (byHand) {
      manual.add(field);
      guessed.delete(field);
    } else if (guessedFields.includes(field)) {
      guessed.add(field);
    } else {
      guessed.delete(field);
    }
  }

  next.guessed = [...guessed];
  if (byHand) next.manualFields = [...manual];

  // Only where the record already keeps them. A draft row carries the
  // codes it could not place; a saved asset has no such column, and
  // inventing one here would write a field SharePoint does not have.
  if (Array.isArray(record.additionalCodes) && additional.length) {
    const codes = [...record.additionalCodes];
    for (const value of additional) if (!codes.includes(value)) codes.push(value);
    next.additionalCodes = codes;
  }

  return { record: next, heldBack };
}
