/**
 * Bytes as something a person can read.
 *
 * In its own module, not beside the components that use it: a file that
 * exports a component must export nothing else, or it drops out of Fast
 * Refresh and fails `npm run lint`. Same rule that put `initialsOf` in
 * `src/utils/initials.js`.
 */
export function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return '—';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
