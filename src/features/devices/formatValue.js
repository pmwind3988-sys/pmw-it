import { formatMYT } from '../../utils/malaysiaTime.js';

/**
 * One value, rendered the way its kind reads.
 *
 * `false` is the one that bites: React renders a bare boolean as nothing at
 * all, so an unprotected machine would show an empty cell rather than "No".
 */
export function formatScalar(value, kind) {
  if (value === null || value === undefined || value === '') return '—';
  if (kind === 'boolean' || typeof value === 'boolean') return value ? 'Yes' : 'No';
  // A timestamp, not a number to read: `importedOn` would otherwise show as
  // thirteen digits of epoch milliseconds.
  if (kind === 'datetime') return formatMYT(value, 'datetime12');
  if (Array.isArray(value)) return value.length ? value.join(', ') : '—';
  return String(value);
}
