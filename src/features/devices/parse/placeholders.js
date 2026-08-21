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
