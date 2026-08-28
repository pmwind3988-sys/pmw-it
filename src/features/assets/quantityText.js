/**
 * What a half-typed quantity box means.
 *
 * An empty box is not a row that owns nothing — it is somebody two keystrokes
 * into replacing 1 with 3. Reading it as a number the moment it is empty is
 * what put a 1 back under the cursor and made "3" come out as "13", which is
 * the difference between one monitor and thirteen.
 *
 * So an empty box commits NOTHING and the row keeps the count it had. The
 * value is only read back when there is something to read.
 */

/**
 * The number to commit for what has been typed, or `null` for "not yet".
 *
 * Zero, a negative and a word are all `null` rather than 1: the box is left
 * showing what was typed while the row keeps its old count, so somebody typing
 * "0" sees their own keystroke instead of a 1 they did not type. What they
 * meant by it is settled when they leave the box.
 */
export function typedQuantity(text) {
  const trimmed = String(text ?? '').trim();
  if (!trimmed) return null;

  const parsed = Number(trimmed);
  if (!Number.isFinite(parsed) || parsed < 1) return null;

  return Math.floor(parsed);
}

/** What the box shows once it is left: the typed number, or the old count back. */
export function settledQuantity(text, previous) {
  const typed = typedQuantity(text);
  if (typed !== null) return typed;

  const before = Math.trunc(Number(previous));
  return Number.isFinite(before) && before >= 1 ? before : 1;
}
