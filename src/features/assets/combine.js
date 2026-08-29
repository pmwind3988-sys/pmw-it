import { BULK } from './assetKinds.js';
import { placeUnit, parseUnits, serialiseUnits } from './units.js';
import { isOpen, out } from './handover/availability.js';

/**
 * Ten monitors that should have been one line of ten.
 *
 * A delivery scanned before the "more than one of something is a bulk line"
 * rule landed — or scanned one box at a time, each with its own serial —
 * arrives in the register as ten rows. Every count on every page then reads
 * ten separate purchases of one monitor, the pager has nothing to page, and
 * handing one out means picking which of ten identical rows to hand out from.
 *
 * This puts them back together: one line with a quantity of ten, and each of
 * the ten rows' serial, label, condition and status kept as ONE ITEM on it.
 * Nothing typed about any of them is lost — that is the whole point, and it is
 * why the row-level fields are moved into item records rather than dropped
 * with the rows.
 *
 * Pure. The write it describes is in `sharepoint/combineAssets.js`.
 */

/** What is on the row rather than on its items, and would go if the row did. */
const SAME_THING = [
  { key: 'category', label: 'category' },
  { key: 'manufacturer', label: 'make' },
  { key: 'model', label: 'model' },
];

const countOf = (row) => Math.max(1, Math.trunc(Number(row?.quantity) || 1));

/**
 * Whether this handover came off this row.
 *
 * The row id is preferred over the key because two bulk rows of the same thing
 * SHARE a key — `bulk:MONITOR|DELL|P2422H` names both of them — and those are
 * exactly the rows somebody reaches for this feature to combine. Falling back
 * to the key covers a handover written before the row id was recorded.
 */
function holds(handover, row) {
  const id = Number(handover?.assetId);
  if (Number.isFinite(id) && id) return id === row?.id;
  return Boolean(handover?.assetKey) && handover.assetKey === row?.assetKey;
}

/**
 * Which item of its row a handover names, or `null` for one that names none.
 *
 * Guarded rather than trusted, because a missing unit reads as `null` and
 * `Number(null)` is 0 — which would quietly turn "a whole laptop" into
 * "item 1 of it".
 */
function unitAt(handover) {
  const value = handover?.unitIndex;
  if (value === null || value === undefined || value === '') return null;
  const index = Number(value);
  return Number.isInteger(index) && index >= 0 ? index : null;
}

/**
 * One row's item records moved into its own block of the combined line.
 *
 * Item 2 of a row stays item 2 OF THAT ROW: it lands at `at + 2`, not at the
 * next free slot. Nothing would be lost by packing them together — every
 * record survives either way — but a handover can name an item nobody has
 * written a serial on yet, and then its NUMBER is the only thing identifying
 * it. Keeping the numbers is what lets one rule place the records and the
 * people holding them.
 */
function placeBlock(stored, row, at) {
  const base = parseUnits(stored);
  const taken = new Set(base.map((unit) => unit.index));
  const moved = new Map();
  const added = [];

  for (const unit of parseUnits(row?.units)) {
    // Stepped over rather than written on, on the off chance a row carries an
    // item numbered past its own quantity — a quantity lowered after the fact
    // leaves those behind, and they describe real objects too.
    let index = at + unit.index;
    while (taken.has(index)) index += 1;
    taken.add(index);
    moved.set(unit.index, index);
    added.push({ ...unit, index });
  }

  return { units: serialiseUnits([...base, ...added]), moved };
}

/** Oldest first: the row everything else is being folded into is the original. */
function byOldest(a, b) {
  const arrived = (Number(a?.arrivedOn) || 0) - (Number(b?.arrivedOn) || 0);
  if (arrived) return arrived;
  return (Number(a?.id) || 0) - (Number(b?.id) || 0);
}

/**
 * Why these rows cannot be combined, or an empty list.
 *
 * Being out with somebody is deliberately NOT one of them. A monitor on
 * somebody's desk is still one of the ten this line was always meant to be,
 * and making people fetch it back before the register can be tidied is asking
 * them to move furniture to fix a typo. The handover moves with the item
 * instead: it keeps naming the same person and the same serial, and points at
 * the item's new position on the combined line (`stillOut`, and the repoint in
 * `sharepoint/combineAssets.js`).
 */
export function blockersFor(rows) {
  if (!Array.isArray(rows) || rows.length < 2) return ['Pick two rows or more.'];
  return [];
}

/**
 * The handovers on these rows that are still open — what somebody is holding
 * right now, so the screen can say plainly that combining will not disturb it.
 */
export function stillOut(rows, handovers = []) {
  return (handovers ?? []).filter(
    (handover) => isOpen(handover) && (rows ?? []).some((row) => holds(handover, row)),
  );
}

/** What differs between the rows, said plainly, so a wrong pick is spotted. */
export function differencesIn(rows) {
  return SAME_THING
    .filter(({ key }) => new Set(
      rows.map((row) => String(row?.[key] ?? '').trim().toLowerCase()),
    ).size > 1)
    .map(({ label }) => label);
}

/**
 * The one line these rows should have been, and the rows it replaces.
 *
 * The keeper is the oldest row: its change history, its delivery and its
 * photographs are already attached to it, and folding the older into the
 * newer would throw the longer record away.
 *
 * Each row contributes its own item records first and then whatever the row
 * itself said about one physical thing — a tracked row's serial number is that
 * row's ONE item, and it takes the next free position rather than overwriting
 * item 1. Positions are handed out a row at a time, so a line of three
 * followed by a line of two occupies items 1-3 and 4-5.
 *
 * `moves` is the other half of that, and it is what lets a row that is out
 * with somebody be combined at all. A handover names the row it came from and
 * which item of it; both of those are about to change, so each handover on
 * every one of these rows — the keeper's included, because a tracked row
 * becoming a bulk line changes its key too — is listed here with the item
 * number its thing has just been given. Returned handovers are listed as well:
 * "who had this monitor last year" is worth as much as "who has it now", and
 * it would be pointing at a row that no longer exists.
 */
export function planCombine(rows, handovers = []) {
  const ordered = [...rows].sort(byOldest);
  const [keep, ...remove] = ordered;

  let units = '';
  let at = 0;
  const moves = [];

  for (const row of ordered) {
    const stored = placeBlock(units, row, at);
    units = stored.units;

    // The row's own serial, label, condition and status as ONE item.
    const own = placeUnit(units, row, at);
    units = own.units;

    for (const handover of handovers) {
      if (!holds(handover, row)) continue;
      const was = unitAt(handover);
      moves.push({
        // A handover naming an item of a bulk row follows that item to its new
        // number. One naming no item at all — a whole tracked row, or a plain
        // count off a bulk line — takes the position the ROW itself just became,
        // which for a tracked row is the item wearing its serial.
        was: handover,
        row,
        unitIndex: was === null ? own.index : (stored.moved.get(was) ?? at + was),
      });
    }

    at += countOf(row);
  }

  return {
    keep,
    remove,
    moves,
    // Only what has to change. Everything else on the keeper — where it came
    // from, its supplier, its PO, its photographs — is already right.
    edits: {
      trackingMode: BULK,
      quantity: at,
      units,
      // What is with people, added up. The count of what the company OWNS and
      // the count of what is OUT have to move together, or nine monitors on
      // nine desks come back onto the shelf the moment their rows are folded
      // in, and the register offers to hand them out again.
      quantityOut: ordered.reduce((sum, row) => sum + out(row), 0),
      // A bulk line can be with five people at once, so it never names one
      // holder (§4.2). The tracked row's copy of who has it goes, and the
      // handover list — which is the truth — keeps every one of them.
      assignedTo: '',
      assignedToEmail: '',
    },
    warnings: differencesIn(ordered),
    recorded: parseUnits(units).length,
  };
}
