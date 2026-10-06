import { normaliseCode } from '../identity.js';

/**
 * The camera reports a code on every frame it can see it — a box held still
 * for a second is thirty reads. The till counts a repeat scan UP, so without
 * this one box of mice would become thirty.
 *
 * A code counts again only after it has been out of sight for `gapMs`. Every
 * sighting pushes the window on, so a box held in frame never re-counts, and
 * taking it away and bringing it back is a deliberate second scan.
 */
export function createCooldown(gapMs = 1200) {
  const lastSeen = new Map();

  return function fresh(raw, now = Date.now()) {
    const code = normaliseCode(raw);
    if (!code) return false;
    const previous = lastSeen.get(code);
    lastSeen.set(code, now);
    return previous === undefined || now - previous > gapMs;
  };
}
