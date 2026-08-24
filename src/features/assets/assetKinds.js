/**
 * What kind of thing IT owns, and whether it is counted one at a time or by
 * the bagful.
 *
 * A TRACKED item is one physical unit with its own serial number, photo,
 * sticker label and history. A BULK line is a model with a quantity that goes
 * up and down — twenty identical mice are one row reading 20, not twenty rows.
 *
 * The kind follows from the category rather than being asked separately, so
 * two people cannot record the same sort of thing two different ways. Where
 * the mapping is wrong for a particular purchase — a serialised keyboard, a
 * spare monitor nobody wants to track — the review grid can override it on the
 * row; this is the default, not a law.
 */

export const TRACKED = 'Tracked';
export const BULK = 'Bulk';

/**
 * Order matters: this is the order the category dropdown offers, so the things
 * bought most often sit at the top rather than in alphabetical exile.
 */
export const CATEGORIES = [
  'Laptop',
  'Desktop',
  'Monitor',
  'Printer',
  'Docking Station',
  'Phone',
  'Tab',
  'Network',
  'PC Part',
  'Keyboard',
  'Mouse',
  'Cable',
  'Adapter',
  'Accessory',
  'Other',
];

/**
 * Everything with a serial number worth keeping. `PC Part` is here because a
 * stick of RAM or a spare SSD carries a serial and is worth following into the
 * machine it ends up in; `Accessory` is not, because it is the catch-all for
 * bags and stands.
 *
 * `Tab` is deliberately absent. Tablets arrive as a box of identical units and
 * are bought that way — two tabs are one row reading 2, not two rows. Their
 * individual serials are not lost by that: a bulk row carries a unit record
 * per physical item (`units.js`), which is where one tab's own serial, label
 * and condition live. A purchase that wants a row each still says so on the
 * row, the same as any other category.
 */
const TRACKED_CATEGORIES = new Set([
  'Laptop',
  'Desktop',
  'Monitor',
  'Printer',
  'Docking Station',
  'Phone',
  'Network',
  'PC Part',
]);

export const CONDITIONS = ['New', 'Good', 'Fair', 'Faulty', 'Retired'];

export const STATUSES = [
  'In stock',
  'Assigned',
  'Borrowed',
  'In repair',
  'Retired',
  'Disposed',
];

/** An unknown category is bulk: quantity is the safer default to be wrong in. */
export function trackingModeFor(category) {
  return TRACKED_CATEGORIES.has(category) ? TRACKED : BULK;
}

export function isTracked(mode) {
  return mode === TRACKED;
}
