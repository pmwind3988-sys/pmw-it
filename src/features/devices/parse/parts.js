import { cleanValue } from './placeholders.js';

/**
 * A value line of a scan version 2 block is `cell | cell | cell`. Every cell is
 * trimmed and a blank cell stays in place, because the position of a cell is
 * what says what it is (`Publisher |` at the end is an empty publisher, not a
 * missing one).
 */
export const splitParts = (line) => String(line ?? '').split('|').map((cell) => cell.trim());

/** The scan's way of saying it looked and found nothing. */
export const isNone = (line) => /^none( found)?$/i.test(String(line ?? '').trim());

/**
 * The value of the first cell that starts with `prefix` (a RegExp that eats the
 * label, e.g. the serial pattern written as a regex literal), cleaned by `clean`. Null when no cell has it, or
 * when the value is a placeholder.
 */
export function cellValue(cells, prefix, clean = cleanValue) {
  const cell = cells.find((entry) => prefix.test(entry));
  return cell ? clean(cell.replace(prefix, '')) : null;
}

/** Lines of a stored multi-line value, without the blanks. */
export const linesOf = (raw) => (raw
  ? String(raw).split('\n').map((line) => line.trim()).filter(Boolean)
  : []);

/** The report lines of a block, with the scan's "nothing found" lines dropped. */
export const meaningful = (lines) => (lines ?? []).filter((line) => !isNone(line));
