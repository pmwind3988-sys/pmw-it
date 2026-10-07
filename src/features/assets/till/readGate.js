import { normaliseCode } from '../identity.js';
import { serialScore, partScore } from '../scan/classifyCode.js';

/**
 * Which barcode the camera is actually being pointed at, and whether to
 * believe it.
 *
 * A box carries four or five barcodes and the camera reads whichever it sees,
 * sometimes wrongly. So nothing a single frame says is taken at face value:
 *
 *  1. READ TWICE. A code counts once it has been read twice within `windowMs`.
 *     A misread is almost always one bad frame, and does not repeat.
 *  2. AIM WINS. The scanner alternates between the aiming box and the whole
 *     picture. A code seen inside the box outranks everything else in view,
 *     so a shipping label at the edge of the frame does not get a say.
 *  3. ASK WHEN UNSURE. Two confirmed codes in the same tier is a question
 *     only the person holding the box can answer, so they are OFFERED as a
 *     choice instead of one being guessed. A code still waiting for its
 *     second read holds the decision back, briefly, rather than letting its
 *     rival win by being faster.
 *  4. ONE ANSWER PER SIGHTING. A code is taken (or offered) once, and again
 *     only after it has been out of sight for `gapMs` — which is also what
 *     stops a box held still from counting thirty times.
 *
 * `expect: 'serial'` settles a choice without asking when exactly one of the
 * offered codes looks like a serial — the till knows it is waiting for one.
 */
export function createReadGate({ windowMs = 1500, gapMs = 1200, need = 2 } = {}) {
  const codes = new Map();

  const forget = (now) => {
    for (const [code, entry] of codes) {
      entry.hits = entry.hits.filter((hit) => now - hit.at <= windowMs);
      if (now - entry.lastSeen > gapMs) {
        if (!entry.hits.length) codes.delete(code);
        else entry.used = false;
      }
    }
  };

  /**
   * One frame's codes. Returns `{ accept: string[], choices: string[] }`:
   * codes to act on now, and — when the gate cannot tell — codes to ask about.
   */
  return function feed(frameCodes = [], { aimed = false, now = Date.now(), expect = null } = {}) {
    forget(now);

    for (const raw of new Set(frameCodes.map((entry) => normaliseCode(entry?.rawValue ?? entry)).filter(Boolean))) {
      const entry = codes.get(raw) ?? { hits: [], used: false, lastSeen: 0 };
      if (now - entry.lastSeen > gapMs) entry.used = false;
      entry.hits.push({ at: now, aimed });
      entry.lastSeen = now;
      codes.set(raw, entry);
    }

    const live = [...codes].filter(([, entry]) => !entry.used && entry.hits.length);
    const isAimed = ([, entry]) => entry.hits.some((hit) => hit.aimed);
    // While ANYTHING is being read inside the aiming box — even a code already
    // taken — codes outside it stay silent. Otherwise the shipping label at
    // the edge of the picture wins the moment the aimed code has been used.
    const aiming = [...codes].some(isAimed);
    const tier = aiming ? live.filter(isAimed) : live;
    const confirmed = tier.filter(([, entry]) => entry.hits.length >= need);

    if (!confirmed.length) return { accept: [], choices: [] };

    const waiting = tier.length - confirmed.length;
    const mark = (list) => { for (const [, entry] of list) entry.used = true; };

    if (confirmed.length === 1) {
      if (waiting) return { accept: [], choices: [] };
      mark(confirmed);
      return { accept: [confirmed[0][0]], choices: [] };
    }

    if (expect === 'serial') {
      const serials = confirmed.filter(([code]) => looksLikeSerial(code));
      if (serials.length === 1) {
        mark(confirmed);
        return { accept: [serials[0][0]], choices: [] };
      }
    }

    mark(confirmed);
    return { accept: [], choices: confirmed.map(([code]) => code) };
  };
}

export function looksLikeSerial(code) {
  return serialScore(code) > partScore(code);
}

/** A few words on what a code probably is, for the person choosing. */
export function guessKind(code) {
  const value = normaliseCode(code);
  if (/^\d{12,14}$/.test(value)) return 'Shop barcode — usually the model';
  if (/^([0-9A-F]{2}[:-]){5}[0-9A-F]{2}$/.test(value) || /^[0-9A-F]{12}$/.test(value)) return 'Looks like a MAC address';
  return looksLikeSerial(value) ? 'Looks like a serial number' : 'Looks like a part or model number';
}
