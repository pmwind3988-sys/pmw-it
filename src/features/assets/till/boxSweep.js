import { MAKES, readTextFields, MIN_CONFIDENCE } from '../scan/classifyText.js';

/**
 * A box sweep: the camera stays on a box while it is turned, every side is
 * read, and what the printing says becomes suggestions — the make, the model,
 * what kind of thing it is, its colour and the details worth a line in the
 * register.
 *
 * Nothing here is believed on one sighting. A value counts as read once it has
 * been seen in two passes (`AGREE`), the same rule the label reader uses, and
 * the most-seen value of each kind wins. Everything it offers is a GUESS: it
 * is shown as one, switched off with a tap, and only written when the person
 * says so.
 *
 * It reads only what is printed. A stylised logo, a curved glossy face or
 * a tiny font is missed, and a colour is never guessed from the artwork —
 * a black mouse is often photographed on a red box.
 */

export const AGREE = 2;

const MAKE_NAMES = [...MAKES].sort((a, b) => b.length - a.length);

/** Words that say what a thing IS, mapped to the register's categories. */
const KIND_WORDS = [
  [/\b(?:docking\s*station|dock)\b/i, 'Docking Station'],
  [/\b(?:laptop|notebook|ultrabook|chromebook)\b/i, 'Laptop'],
  [/\b(?:desktop|mini\s*pc|tower)\b/i, 'Desktop'],
  [/\b(?:monitor|display)\b/i, 'Monitor'],
  [/\b(?:printer|laserjet|inkjet|deskjet|officejet)\b/i, 'Printer'],
  [/\b(?:smartphone|phone)\b/i, 'Phone'],
  [/\b(?:tablet|ipad)\b/i, 'Tab'],
  [/\b(?:keyboard)\b/i, 'Keyboard'],
  [/\b(?:mouse|mice)\b/i, 'Mouse'],
  [/\b(?:cable|cord)\b/i, 'Cable'],
  [/\b(?:adapter|adaptor|dongle|converter|hub)\b/i, 'Adapter'],
  [/\b(?:router|switch|access\s*point)\b/i, 'Network'],
  [/\b(?:headset|headphones?|earphones?|webcam|speaker|charger)\b/i, 'Accessory'],
];

const COLOURS = [
  'Space Grey', 'Space Gray', 'Rose Gold', 'Black', 'White', 'Silver', 'Grey', 'Gray',
  'Graphite', 'Blue', 'Navy', 'Red', 'Pink', 'Gold', 'Green', 'Purple', 'Beige',
];

/** Details worth a line in the register, each read by its own shape. */
const DETAIL_PATTERNS = [
  ['connection', /\b(?:wireless|bluetooth|wired|2\.4\s*ghz)\b/i],
  ['connector', /\b(?:usb[-\s]?c|usb\s*type[-\s]?c|usb[-\s]?a|usb\s*3(?:\.\d)?|usb\s*2\.0|hdmi|displayport|thunderbolt\s*\d?|lightning|rj45)\b/i],
  ['length', /\b\d+(?:\.\d+)?\s?(?:m|metres?|meters?|ft)\b/i],
  ['capacity', /\b\d+(?:\.\d+)?\s?(?:GB|TB)\b/i],
  ['screen', /\b\d{2}(?:\.\d)?\s?(?:"|”|inch|in\b|-inch)/i],
  ['resolution', /\b(?:FHD|UHD|QHD|4K|\d{3,4}\s?x\s?\d{3,4})\b/i],
  ['power', /\b\d{2,3}\s?W\b/],
];

const DETAIL_ORDER = ['connection', 'connector', 'length', 'capacity', 'screen', 'resolution', 'power'];

/** Words a model name on a box front trails off into, which are not the model. */
const GENERIC = /\b(?:wireless|wired|bluetooth|optical|usb|mouse|mice|keyboard|monitor|display|headset|cable|adapter|combo|laptop|notebook|webcam|for\s+.*)\b.*$/i;

const tidy = (value) => String(value ?? '').replace(/\s+/g, ' ').replace(/[®™©]/g, '').trim();

function titleCase(word) {
  return word.replace(/\b\w/g, (c) => c.toUpperCase()).replace(/\b(Usb|Hdmi|Fhd|Uhd|Qhd|Ghz|Gb|Tb|Rj45)\b/g, (m) => m.toUpperCase());
}

/**
 * "Logitech M90 Wireless Mouse" — a known make starting the line, and the
 * model after it, up to the first generic word.
 */
export function makeAndModel(line) {
  const text = tidy(line);
  for (const make of MAKE_NAMES) {
    const at = text.toLowerCase().indexOf(make.toLowerCase());
    if (at !== 0) continue;
    const rest = tidy(text.slice(make.length))
      .replace(GENERIC, '')
      // A screen size printed after the model is the size, not the model.
      .replace(/\s+\d{2}(?:\.\d)?\s?(?:"|”|-?inch).*$/i, '')
      .trim();
    // A real model has a digit or is short and specific; "Logitech Advanced
    // Optical Tracking" is a slogan.
    const model = /\d/.test(rest) && rest.length <= 24 ? rest : '';
    return { make, model };
  }
  return null;
}

/** Everything one line of box text says, as votes. */
export function readLine(line) {
  const text = tidy(line);
  const votes = [];
  const named = makeAndModel(text);
  if (named) {
    votes.push(['make', named.make]);
    if (named.model) votes.push(['model', named.model]);
  }
  for (const [pattern, category] of KIND_WORDS) {
    if (pattern.test(text)) { votes.push(['category', category]); break; }
  }
  for (const colour of COLOURS) {
    if (new RegExp(`\\b${colour}\\b`, 'i').test(text)) {
      votes.push(['colour', colour === 'Gray' ? 'Grey' : colour.replace('Space Gray', 'Space Grey')]);
      break;
    }
  }
  for (const [kind, pattern] of DETAIL_PATTERNS) {
    const found = text.match(pattern);
    if (!found) continue;
    const raw = tidy(found[0]);
    // Lengths keep their small units: "1.8 m", never "1.8 M".
    votes.push([kind, kind === 'length' ? raw.replace(/\s?(m|ft)\b/i, (_, unit) => ` ${unit.toLowerCase()}`) : titleCase(raw)]);
  }
  return votes;
}

export function newSweep() {
  return { passes: 0, votes: {} };
}

/**
 * One pass of the camera: the lines it read. A value is counted once per
 * pass however many lines repeat it, so one side printed five times does not
 * outvote another side read twice.
 */
export function recordSweep(sweep, lines = []) {
  // A line the reader itself was unsure of is a smudge or a logo, not text.
  const texts = lines
    .filter((line) => typeof line === 'string' || !(typeof line?.confidence === 'number' && line.confidence < MIN_CONFIDENCE))
    .map((line) => (typeof line === 'string' ? line : line?.text))
    .filter(Boolean);
  const seen = new Set();
  const add = (kind, value) => {
    const v = tidy(value);
    if (!v) return;
    seen.add(`${kind}\u0000${v}`);
  };

  for (const text of texts) for (const [kind, value] of readLine(text)) add(kind, value);

  // The label reader's own answers too — a sticker on the box saying
  // "Model: M90" is the strongest evidence a box has.
  const fields = readTextFields(texts.map((text) => ({ text })));
  if (fields.manufacturer) add('make', fields.manufacturer);
  if (fields.model) add('model', fields.model);

  const votes = { ...sweep.votes };
  for (const entry of seen) {
    const [kind, value] = entry.split('\u0000');
    const tally = { ...(votes[kind] ?? {}) };
    tally[value] = (tally[value] ?? 0) + 1;
    votes[kind] = tally;
  }
  return { passes: sweep.passes + 1, votes };
}

function best(tally = {}, agree = AGREE) {
  const ranked = Object.entries(tally).filter(([, n]) => n >= agree).sort((a, b) => b[1] - a[1]);
  return ranked.length ? ranked[0][0] : '';
}

/** What the sweep is sure enough of to offer. */
export function sweepSuggestions(sweep, agree = AGREE) {
  const details = DETAIL_ORDER.map((kind) => best(sweep.votes[kind], agree)).filter(Boolean);
  return {
    category: best(sweep.votes.category, agree),
    make: best(sweep.votes.make, agree),
    model: best(sweep.votes.model, agree),
    colour: best(sweep.votes.colour, agree),
    details,
  };
}

/** Values seen once so far — shown as "reading…" so the person keeps turning. */
export function sweepPending(sweep, agree = AGREE) {
  const out = [];
  for (const [kind, tally] of Object.entries(sweep.votes)) {
    for (const [value, n] of Object.entries(tally)) if (n < agree) out.push({ kind, value });
  }
  return out;
}

/** The colour and details as one line for the register's details field. */
export function detailLine({ colour = '', details = [] } = {}) {
  return [colour, ...details].filter(Boolean).join(' · ');
}
