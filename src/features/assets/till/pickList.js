import { normaliseCode } from '../identity.js';

/**
 * What the full-screen till camera offers to be tapped.
 *
 * The camera no longer decides which barcode on a box was meant: it lists
 * what it reads and the person holding the box taps the ones they want.
 * The rules here keep that list honest:
 *
 *  - a code appears only once it has been read twice within `windowMs` — a
 *    misread is one bad frame and does not repeat;
 *  - a code the till would refuse as already listed (`skip`) never appears;
 *  - a tapped code stays on the card as "added" while it is in view and
 *    cannot be tapped twice;
 *  - a code is forgotten `keepMs` after it was last seen — the card clears
 *    itself, and the next identical box of mice can be tapped again;
 *  - at most `max` codes, oldest first, so a row does not move under a finger.
 *
 * Pure: the camera calls `seePicks` per frame and `visiblePicks` to draw.
 */
export const PICK_TIMING = { need: 2, windowMs: 1500, keepMs: 4000, max: 6 };

export function newPicks() {
  return {};
}

const forget = (picks, now) => Object.fromEntries(
  Object.entries(picks).filter(([, entry]) => now - entry.lastSeen <= PICK_TIMING.keepMs),
);

export function seePicks(picks, frameCodes = [], { now = Date.now(), skip = () => false } = {}) {
  const next = forget(picks, now);
  const codes = new Set(frameCodes.map((entry) => normaliseCode(entry?.rawValue ?? entry)).filter(Boolean));
  for (const code of codes) {
    // A code tapped here is on the list now, so `skip` says yes to it — but
    // it is shown as "added", not hidden, so the tap visibly did something.
    if (!next[code]?.taken && skip(code)) {
      delete next[code];
      continue;
    }
    const entry = next[code] ?? { firstSeen: now, hits: [], shown: false, taken: false };
    const hits = [...entry.hits.filter((at) => now - at <= PICK_TIMING.windowMs), now];
    next[code] = {
      ...entry,
      hits,
      lastSeen: now,
      shown: entry.shown || hits.length >= PICK_TIMING.need,
    };
  }
  return next;
}

export function takePick(picks, code) {
  const key = normaliseCode(code);
  if (!picks[key]) return picks;
  return { ...picks, [key]: { ...picks[key], taken: true } };
}

export function visiblePicks(picks, now = Date.now()) {
  return Object.entries(forget(picks, now))
    .filter(([, entry]) => entry.shown)
    .sort(([, a], [, b]) => a.firstSeen - b.firstSeen)
    .slice(-PICK_TIMING.max)
    .map(([code, entry]) => ({ code, taken: entry.taken }));
}
