import { BULK } from './assetKinds.js';
import { mergeUnits, appendUnit, parseUnits } from './units.js';
import { isOpen } from './handover/availability.js';

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

/** Oldest first: the row everything else is being folded into is the original. */
function byOldest(a, b) {
  const arrived = (Number(a?.arrivedOn) || 0) - (Number(b?.arrivedOn) || 0);
  if (arrived) return arrived;
  return (Number(a?.id) || 0) - (Number(b?.id) || 0);
}

/**
 * Why these rows cannot be combined, or an empty list.
 *
 * An open handover is a blocker rather than a warning. A handover record names
 * the row it came from; folding that row away would leave somebody holding an
 * item the register cannot account for, and "who has what" would quietly stop
 * matching the shelf. Bring the thing back first, then combine.
 */
export function blockersFor(rows, handovers = []) {
  const blockers = [];
  if (!Array.isArray(rows) || rows.length < 2) {
    blockers.push('Pick two rows or more.');
    return blockers;
  }

  const out = rows.filter((row) => handovers.some(
    (handover) => handover.assetKey === row.assetKey && isOpen(handover),
  ));

  if (out.length) {
    blockers.push(
      `${out.length === 1 ? 'One of them is' : `${out.length} of them are`} out with `
      + 'somebody. Take those back first, or leave them out of this.',
    );
  }

  return blockers;
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
 */
export function planCombine(rows) {
  const ordered = [...rows].sort(byOldest);
  const [keep, ...remove] = ordered;

  let units = '';
  let at = 0;
  for (const row of ordered) {
    units = mergeUnits(units, row.units, at);
    units = appendUnit(units, row, at);
    at += countOf(row);
  }

  return {
    keep,
    remove,
    // Only what has to change. Everything else on the keeper — where it came
    // from, its supplier, its PO, its photographs — is already right.
    edits: { trackingMode: BULK, quantity: at, units },
    warnings: differencesIn(ordered),
    recorded: parseUnits(units).length,
  };
}
