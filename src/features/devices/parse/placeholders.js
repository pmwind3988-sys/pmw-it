/**
 * Tokens the scan writes when it could not read a real value. Storing them
 * verbatim would produce a dashboard category called "Manufacturer1", so they
 * become null at the parse boundary rather than being filtered downstream.
 */
const PLACEHOLDER_TOKENS = new Set([
  'none',
  'unknown',
  'n/a',
  'na',
  'nil',
  'manufacturer1',
  'partnum1',
  'system product name',
  'to be filled by o.e.m.',
  'default string',
  'not specified',
  '',
]);

/** Non-breaking space, zero-width space, zero-width non-joiner, BOM. */
const INVISIBLE = /[\u00a0\u200b\u200c\ufeff]/g;

export function isPlaceholder(value) {
  if (value == null) return true;
  return PLACEHOLDER_TOKENS.has(String(value).replace(INVISIBLE, ' ').trim().toLowerCase());
}

export function cleanValue(value) {
  if (value == null) return null;
  const cleaned = String(value).replace(INVISIBLE, ' ').replace(/\s+$/, '').trim();
  return isPlaceholder(cleaned) ? null : cleaned;
}

/**
 * What a BIOS writes when nobody flashed a real serial in. A home-built
 * desktop reports one of these, and every such desktop reports the SAME one,
 * so matching on it would merge every unbranded PC in the company into one row.
 */
const PLACEHOLDER_SERIALS = new Set([
  'system serial number',
  'chassis serial number',
  'default string',
  'to be filled by o.e.m.',
  'not specified',
  '0',
]);

/** Serials compare upper-cased with every space gone: `cn0abc 123` is `CN0ABC123`. */
export function normaliseSerial(value) {
  return String(value ?? '').replace(INVISIBLE, '').replace(/\s+/g, '').toUpperCase();
}

export function isPlaceholderSerial(value) {
  if (isPlaceholder(value)) return true;
  const spaced = String(value).replace(INVISIBLE, ' ').trim().toLowerCase();
  if (PLACEHOLDER_SERIALS.has(spaced)) return true;
  // `00000000`, `XXXXXXXX`: one character repeated is a filler, not a serial.
  return /^(.)\1*$/.test(normaliseSerial(value));
}

/** The serial as stored, or null when the scan had nothing usable. */
export function cleanSerial(value) {
  return isPlaceholderSerial(value) ? null : normaliseSerial(value);
}
