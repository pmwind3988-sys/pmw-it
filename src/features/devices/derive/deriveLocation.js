import { BUILT_IN_LOCATIONS, cleanLocation } from '../map/locations.js';

/**
 * The location IT writes first inside the file name's bracket:
 * `[F1 ENGINEERING] AMIR-HP.txt`. Anything that is not a known code is left
 * alone, so `[ENGINEERING] AMIR-HP.txt` reads exactly as it always did.
 *
 * Two names predate the convention and carry the site glued on or in front:
 * `STOCKYARDF1` and `PML GUARDHOUSE`. A glued code is split off only when it is
 * a KNOWN code with at least three letters in front of it, so a word that
 * merely ends in something code-shaped is not cut.
 */
export function splitLocation(bracket, known = BUILT_IN_LOCATIONS) {
  if (!bracket || !bracket.trim()) return { location: null, rest: null };

  const codes = [...new Set(known.map(cleanLocation).filter(Boolean))]
    .sort((a, b) => b.length - a.length);
  const words = bracket.trim().split(/\s+/);
  const first = words[0].toUpperCase();
  const tail = words.slice(1).join(' ') || null;

  if (codes.includes(first)) return { location: first, rest: tail };

  for (const code of codes) {
    if (first.length < code.length + 3 || !first.endsWith(code)) continue;
    const head = words[0].slice(0, -code.length);
    if (!/[A-Z]$/i.test(head)) continue;
    return { location: code, rest: [head, tail].filter(Boolean).join(' ') };
  }

  return { location: null, rest: bracket.trim() };
}
