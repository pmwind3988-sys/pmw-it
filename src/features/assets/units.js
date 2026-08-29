import { CONDITIONS, STATUSES, TRACKED } from './assetKinds.js';

/**
 * The individual things inside a bulk line.
 *
 * A bulk row says "2 tabs". That is the right answer to "what did we buy" and
 * the wrong answer to "which one is Aisyah holding, and which one has the
 * cracked screen". A unit record fills that gap: one entry per physical item
 * on the row, without splitting the line into two rows and losing the fact
 * that it was one purchase of two identical things.
 *
 * A serial number, a part number, a MAC address, a sticker label, a condition
 * and a status all describe ONE physical thing. A bulk row therefore does not
 * hold them at all — `PER_UNIT_ONLY` is the list, and the row is stripped of
 * them wherever it is written. A row carrying one serial for twenty items is
 * not a record of twenty items; it is a record of one, with nineteen hidden
 * behind it.
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
  { key: 'partNumber', label: 'Part number' },
  { key: 'macAddress', label: 'MAC address' },
  { key: 'assetTag', label: 'Asset label' },
  { key: 'condition', label: 'Condition', options: CONDITIONS },
  { key: 'status', label: 'Status', options: STATUSES },
  { key: 'location', label: 'Where it is' },
  { key: 'remarks', label: 'Remarks', multiline: true },
];

/**
 * What a BULK row must never hold, because each names one physical item.
 *
 * `location` and `remarks` are absent from this list on purpose: a line can
 * honestly live in one place and carry one note, and a unit may override
 * either. These six cannot be overridden — they are either about one item or
 * about nothing.
 */
export const PER_UNIT_ONLY = [
  'serialNumber', 'partNumber', 'macAddress', 'assetTag', 'condition', 'status',
];

/**
 * The subset a SCAN can honestly claim for one item: the codes it read off the
 * box in front of it.
 *
 * Condition and status are deliberately not here. The condition on a review
 * grid is a blanket statement about the delivery — "all new" — and writing it
 * onto item 1 alone would turn "twenty new cables" into "one new cable and
 * nineteen nobody looked at", which is worse than the honest nothing.
 */
export const PER_UNIT_CODES = ['serialNumber', 'partNumber', 'macAddress', 'assetTag'];

/**
 * The photograph of ONE item, and a photo not yet uploaded.
 *
 * `photoUrl` is where the picture ended up in SharePoint. `photoId` is a
 * just-taken photo still in the phone's own storage, waiting for the next Save
 * to upload it — carried on the unit record so that closing the page does not
 * lose the picture, and replaced by `photoUrl` once the upload succeeds.
 */
export const UNIT_PHOTO_FIELDS = ['photoUrl', 'photoId'];

const KEYS = [...UNIT_FIELDS.map((field) => field.key), ...UNIT_PHOTO_FIELDS];

/**
 * NOT trimmed, and that is the whole point of it.
 *
 * Every keystroke in the pager leaves through `serialiseUnits` and comes back
 * through `parseUnits`, so anything trimmed here is trimmed WHILE SOMEBODY IS
 * TYPING. A space is only ever typed at the end of what is there so far, so a
 * trim on that round trip makes the space bar do nothing at all, in every
 * field on the card. Trimming happens once, on the way to SharePoint, in
 * `trimUnits`.
 */
const asText = (value) => (value === null || value === undefined ? '' : String(value));

const isBlankValue = (value) => asText(value).trim() === '';

/** A unit nobody has written anything on is not worth storing. */
export function isBlankUnit(unit) {
  return KEYS.every((key) => isBlankValue(unit?.[key]));
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
 * The row's own per-unit values, as the unit record they should have been.
 *
 * Rows written before this rule existed carry a serial, a label and a
 * condition on the row — the Lenovo tabs did. Reading them as item 1 is what
 * stops the change from looking like the register quietly lost them, and it
 * happens at READ time: nothing is written until somebody saves.
 */
function legacyUnit(asset, fields = PER_UNIT_ONLY) {
  const source = {};
  for (const field of fields) source[field] = asset?.[field];

  const unit = cleanUnit(source, 0);
  return isBlankUnit(unit) ? null : unit;
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
  const parsed = parseUnits(stored);
  const filled = new Map(parsed.map((unit) => [unit.index, unit]));

  // Only when there is nothing stored at all. A row that already has unit
  // records has been saved under the new rule, and its row-level leftovers —
  // if any survive — are not a first item waiting to be adopted.
  if (!parsed.length) {
    const legacy = legacyUnit(asset);
    if (legacy) filled.set(0, legacy);
  }

  const count = Math.max(1, Math.trunc(Number(asset?.quantity) || 1));

  const units = [];
  for (let index = 0; index < count; index += 1) {
    units.push(filled.get(index) ?? cleanUnit(null, index));
  }
  return units;
}

/**
 * The row as it should be stored: a bulk line with its per-item fields moved
 * into the units where they belong, and blanked on the row itself.
 *
 * Both halves have to happen together. Blanking without moving loses the data;
 * moving without blanking leaves the register showing one item's serial as if
 * it described the whole line, which is the thing being fixed.
 *
 * A tracked row is returned untouched — it IS one item, so the row is the
 * right place for all of it.
 */
export function withUnitsSplitOut(record, moved = PER_UNIT_ONLY) {
  if (record?.trackingMode === TRACKED) return record;

  const next = { ...record };
  const stored = parseUnits(record?.units);

  if (!stored.length) {
    const legacy = legacyUnit(record, moved);
    if (legacy) next.units = serialiseUnits([legacy]);
  }

  // The one place the stray spaces come off. Everything upstream of here is
  // somebody typing, and a value trimmed mid-word is a space bar that does
  // nothing; everything downstream is storage, where " HA2KJDSW " and
  // "HA2KJDSW" must not be two different serial numbers.
  next.units = trimUnits(next.units);

  // Cleared in full whatever was moved: a field left on the row would go on
  // describing the whole line, which is the thing this exists to stop.
  for (const field of PER_UNIT_ONLY) next[field] = '';
  return next;
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

/**
 * The JSON the column holds. Empty string when there is nothing to keep.
 *
 * Blank fields are dropped rather than written as `""`. A single note column
 * holds every unit on the row, and twenty items each carrying eight empty
 * strings is most of a SharePoint text limit spent on nothing.
 */
export function serialiseUnits(units) {
  const kept = (units ?? [])
    .map((unit, position) => cleanUnit(unit, Number.isInteger(unit?.index) ? unit.index : position))
    .filter((unit) => !isBlankUnit(unit))
    .sort((a, b) => a.index - b.index)
    .map((unit) => {
      const compact = { index: unit.index };
      for (const key of KEYS) if (!isBlankValue(unit[key])) compact[key] = unit[key];
      return compact;
    });

  return kept.length ? JSON.stringify(kept) : '';
}

/**
 * The same records with every value trimmed: what actually gets stored.
 *
 * Kept apart from `serialiseUnits` deliberately. Serialising happens on every
 * keystroke in the pager, and trimming there is what stopped the space bar
 * working; this runs once, when a save is on its way to SharePoint.
 */
export function trimUnits(stored) {
  return serialiseUnits(parseUnits(stored).map((unit) => {
    const trimmed = { index: unit.index };
    for (const key of KEYS) trimmed[key] = asText(unit[key]).trim();
    return trimmed;
  }));
}

/**
 * One more physical item on an existing line, taking the next free position.
 *
 * This is a second box of something already in the register being scanned in:
 * the quantity goes up by one and the thing in the scanner's hand becomes the
 * unit at that new position, rather than overwriting item 1's serial with its
 * own — which is what a plain field-level update would do.
 *
 * `at` is where to put it. A position already taken is stepped over instead of
 * being written on, because the row it would overwrite describes a real object
 * somebody has already recorded.
 *
 * Answers with the position it chose as well as the units, because a caller
 * folding a row away has to be able to say WHICH item the thing became — a
 * handover pointing at that row needs the new number, and working it out from
 * the offset would be wrong the moment a taken position was stepped over.
 */
export function placeUnit(stored, source, at = 0, fields = PER_UNIT_ONLY) {
  const units = parseUnits(stored);
  const candidate = legacyUnit(source, fields);
  if (!candidate) return { units: serialiseUnits(units), index: null };

  const taken = new Set(units.map((unit) => unit.index));
  let index = Math.max(0, Math.trunc(Number(at) || 0));
  while (taken.has(index)) index += 1;

  return { units: serialiseUnits([...units, { ...candidate, index }]), index };
}

/** The same, for the callers that only want the units back. */
export function appendUnit(stored, source, at = 0, fields = PER_UNIT_ONLY) {
  return placeUnit(stored, source, at, fields).units;
}

/**
 * The units of an arriving delivery added to the units already on the row.
 *
 * `offset` is where the new ones start — the row's existing quantity, because
 * everything below that is already spoken for. Positions already taken are
 * stepped over rather than written on: each entry describes a real object
 * somebody has recorded, and two of them must never be merged into one item
 * wearing this tab's serial and that tab's label.
 */
export function mergeUnits(stored, incoming, offset = 0) {
  const base = parseUnits(stored);
  const taken = new Set(base.map((unit) => unit.index));
  const added = [];

  let next = Math.max(0, Math.trunc(Number(offset) || 0));
  for (const unit of parseUnits(incoming)) {
    while (taken.has(next)) next += 1;
    taken.add(next);
    added.push({ ...unit, index: next });
    next += 1;
  }

  return serialiseUnits([...base, ...added]);
}

const LOGGED_KEYS = UNIT_FIELDS.map((field) => field.key);

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
    // The photo fields are deliberately absent: a re-photographed item would
    // otherwise file a change-log line holding two library paths, burying the
    // changes somebody actually reads the log for.
    for (const key of LOGGED_KEYS) {
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

/**
 * A field that is per-item, counted across everything the row owns.
 *
 * `[{ value, count }]`, and the arithmetic is the point: a box of twenty with
 * one unit marked Faulty is one faulty and nineteen unrecorded, not one row
 * that is faulty and not twenty faulty cables. Items nobody has said anything
 * about are counted under `unstated` — 'In stock' for a status, because that
 * is what a thing nobody has handed out is; nothing at all for a condition,
 * because a condition nobody recorded is not a condition.
 *
 * A tracked row is one item and answers with its own value.
 */
export function perItem(asset, field, unstated = '') {
  const owned = Math.max(1, Math.trunc(Number(asset?.quantity) || 1));
  const tally = new Map();
  const add = (value, count) => {
    if (!value || count <= 0) return;
    tally.set(value, (tally.get(value) ?? 0) + count);
  };

  if (asset?.trackingMode === TRACKED) {
    add(asset?.[field] || unstated, 1);
    return [...tally].map(([value, count]) => ({ value, count }));
  }

  // Through `unitsOf` rather than `parseUnits`, so a row still carrying its
  // values on the row counts them as item 1 — the same reading the item detail
  // gives it. Two places disagreeing about one row is how a register starts
  // reporting figures nobody can reproduce by opening it.
  let stated = 0;
  for (const unit of unitsOf(asset)) {
    if (!unit[field]) continue;
    add(unit[field], 1);
    stated += 1;
  }

  add(unstated, owned - stated);
  return [...tally].map(([value, count]) => ({ value, count }));
}

/** How many of the row's items are in this state. */
export function countPerItem(asset, field, wanted, unstated = '') {
  return perItem(asset, field, unstated).find((entry) => entry.value === wanted)?.count ?? 0;
}

/** A one-line name for the unit, for the pager's own heading. */
export function unitTitle(unit, asset) {
  return unit?.serialNumber
    || unit?.assetTag
    || `${asset?.model || asset?.category || 'Item'} #${(unit?.index ?? 0) + 1}`;
}
