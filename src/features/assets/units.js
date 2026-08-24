import { CONDITIONS } from './assetKinds.js';

/**
 * The individual things inside a bulk line.
 *
 * A bulk row says "2 tabs". That is the right answer to "what did we buy" and
 * the wrong answer to "which one is Aisyah holding, and which one has the
 * cracked screen". A unit record fills that gap: one entry per physical item
 * on the row, each with its own serial, label, condition and note, without
 * splitting the line into two rows and losing the fact that it was one
 * purchase of two identical things.
 *
 * Everything here is pure and the storage is one JSON string in a single
 * SharePoint column (`Units`). That is deliberate — a second list would need
 * its own provisioning, its own read, and a join on every page that counts
 * anything, to hold data that is only ever read alongside its own row.
 *
 * A unit is SPARSE. Only the ones somebody has actually filled in are stored,
 * keyed by their position on the row, so a box of twenty cables costs nothing
 * until the day one of them is written on.
 */

export const UNIT_FIELDS = [
  { key: 'serialNumber', label: 'Serial number' },
  { key: 'assetTag', label: 'Asset label' },
  { key: 'condition', label: 'Condition', options: CONDITIONS },
  { key: 'location', label: 'Where it is' },
  { key: 'remarks', label: 'Remarks', multiline: true },
];

const KEYS = UNIT_FIELDS.map((field) => field.key);

const asText = (value) => (value === null || value === undefined ? '' : String(value).trim());

/** A unit nobody has written anything on is not worth storing. */
export function isBlankUnit(unit) {
  return KEYS.every((key) => asText(unit?.[key]) === '');
}

function cleanUnit(raw, index) {
  const unit = { index };
  for (const key of KEYS) unit[key] = asText(raw?.[key]);
  return unit;
}

/**
 * What is stored, read back.
 *
 * Accepts the JSON string the column holds, an already-parsed array, or
 * nothing at all. Anything unreadable answers with an empty list rather than
 * throwing: a row whose unit column got mangled must still open.
 */
export function parseUnits(stored) {
  if (!stored) return [];

  let raw = stored;
  if (typeof stored === 'string') {
    try {
      raw = JSON.parse(stored);
    } catch {
      return [];
    }
  }

  if (!Array.isArray(raw)) return [];

  const byIndex = new Map();
  for (const entry of raw) {
    const index = Number(entry?.index);
    if (!Number.isInteger(index) || index < 0) continue;
    const unit = cleanUnit(entry, index);
    if (!isBlankUnit(unit)) byIndex.set(index, unit);
  }

  return [...byIndex.values()].sort((a, b) => a.index - b.index);
}

/**
 * Every unit on the row, in order, blanks included — which is what a pager
 * needs: "unit 2 of 5" has to exist before anybody can type into it.
 *
 * The count follows the row's quantity, so raising a quantity to 3 adds a
 * third card and lowering it to 1 hides the rest. Lowering does NOT delete
 * them; a quantity typed wrong and corrected back must not silently take a
 * serial number with it. They come back when the quantity does, and are only
 * dropped on the next save that writes over them.
 */
export function unitsOf(asset, stored = asset?.units) {
  const filled = new Map(parseUnits(stored).map((unit) => [unit.index, unit]));
  const count = Math.max(1, Math.trunc(Number(asset?.quantity) || 1));

  const units = [];
  for (let index = 0; index < count; index += 1) {
    units.push(filled.get(index) ?? cleanUnit(null, index));
  }
  return units;
}

/** How many of them somebody has actually recorded something about. */
export function filledCount(units) {
  return units.filter((unit) => !isBlankUnit(unit)).length;
}

export function setUnitField(units, index, field, value) {
  return units.map((unit) => (
    unit.index === index ? { ...unit, [field]: value } : unit
  ));
}

/** The JSON the column holds. Empty string when there is nothing to keep. */
export function serialiseUnits(units) {
  const kept = (units ?? [])
    .map((unit, position) => cleanUnit(unit, Number.isInteger(unit?.index) ? unit.index : position))
    .filter((unit) => !isBlankUnit(unit))
    .sort((a, b) => a.index - b.index);

  return kept.length ? JSON.stringify(kept) : '';
}

const labelFor = (key) => UNIT_FIELDS.find((field) => field.key === key)?.label ?? key;

/**
 * What changed, unit by unit and field by field, for the change log.
 *
 * The log records "Unit 2 · Serial number", not a JSON blob. A blob in the
 * history column is the same as no history at all: nobody reads it, and the
 * one question it exists to answer — when did this unit's label change, and
 * who did it — cannot be asked of it.
 */
export function diffUnits(before, after) {
  const was = new Map(parseUnits(before).map((unit) => [unit.index, unit]));
  const now = new Map(parseUnits(after).map((unit) => [unit.index, unit]));
  const indexes = [...new Set([...was.keys(), ...now.keys()])].sort((a, b) => a - b);

  const changes = [];
  for (const index of indexes) {
    for (const key of KEYS) {
      const oldValue = asText(was.get(index)?.[key]);
      const newValue = asText(now.get(index)?.[key]);
      if (oldValue === newValue) continue;

      changes.push({
        fieldName: `Unit ${index + 1} · ${labelFor(key)}`,
        oldValue,
        newValue,
        changeType: oldValue === '' ? 'Added' : (newValue === '' ? 'Removed' : 'Updated'),
      });
    }
  }

  return changes;
}

/** A one-line name for the unit, for the pager's own heading. */
export function unitTitle(unit, asset) {
  return unit?.serialNumber
    || unit?.assetTag
    || `${asset?.model || asset?.category || 'Item'} #${(unit?.index ?? 0) + 1}`;
}
