/**
 * Two letters standing in for a person, for the avatar chips in the nav and in
 * the records table.
 *
 * Its own module rather than an extra export beside a component: a file that
 * exports both a component and a helper drops out of Fast Refresh, so every
 * edit to the shell would full-reload the page instead of hot-swapping it.
 */
export function initialsOf(name) {
  if (!name) return 'U';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
  return name.slice(0, 2).toUpperCase();
}
