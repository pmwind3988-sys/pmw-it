// Which re-cleans have to carry the grid with them.
//
// The worker KEEPS the parsed grid (see `parse.worker.js`), so a
// re-clean normally costs one small plan message each way however large
// the sheet is. Ticking a cleaning step 100k rows deep must not ship
// 100k rows across the boundary again.
//
// That optimisation has one trap, and it is silent. When the main
// thread REPLACES the grid -- which is exactly what happens when the
// text analysis is added to it as five new columns -- the worker is
// still holding the old one. It cleans that, sends back a dataset with
// no analysis columns in it, and every tile charting them renders
// `Column "Severity" is not in this dataset`. Nothing errors; the
// answer is just quietly the previous sheet.
//
// So the rule is: send the grid exactly when it is not the one the
// worker already has, and never otherwise.

/**
 * The grid to attach to the next clean message, or `undefined` to let
 * the worker use the one it kept.
 *
 * Compared by IDENTITY, not by content. Every path that replaces the
 * grid builds a new object, and a deep comparison of 100k rows on every
 * checkbox tick would cost more than the message it is trying to avoid.
 */
export function gridToSend(current, lastSent) {
  if (!current) return undefined;
  return current === lastSent ? undefined : current;
}
