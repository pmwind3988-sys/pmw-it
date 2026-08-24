import { classifyCodes, serialScore, partScore } from './classifyCode.js';

/**
 * Working out what the words on a label are.
 *
 * A barcode arrives as one value that means one thing. A photographed
 * label arrives as twenty lines, most of which mean nothing: safety
 * marks, a country of origin, a website, the printed noise around the
 * text. This decides which of them are worth keeping and which box each
 * kept line belongs in.
 *
 * The identifiers are NOT decided here. Once the noise and the
 * specification lines are out, what is left is handed to
 * `classifyCodes` — the same scorer the barcode scan uses. Two sets of
 * rules for "is this a serial or a part number" would eventually
 * disagree with each other, and the field that disagreed would be the
 * one nobody checked.
 *
 * Everything worked out rather than read is marked as guessed, for the
 * reason the whole scanning flow marks guesses: a heuristic that cannot
 * be corrected is a heuristic that ships wrong data.
 */

/**
 * Below this, the reader is telling us it could not see the line. On a
 * printed label the low-confidence lines are the smudges and the logos,
 * and one misread character in a serial number is worse than a field
 * left empty.
 */
export const MIN_CONFIDENCE = 55;

/** Two characters is a corner mark or an icon, not a value worth keeping. */
const MIN_LENGTH = 3;

/**
 * Makes IT actually buys. A brand on a line of its own is one of the few
 * things on a label that can be recognised outright, and it saves the
 * most typing.
 *
 * Deliberately a list rather than a shape: nothing about the word
 * `Latitude` says whether it is a make, a model or a room. Only a name
 * we already know is safe to file as the make.
 */
const MAKES = [
  'Dell', 'HP', 'Hewlett-Packard', 'Lenovo', 'Apple', 'Asus', 'Acer', 'MSI',
  'Microsoft', 'Samsung', 'LG', 'Logitech', 'Canon', 'Epson', 'Brother',
  'Cisco', 'TP-Link', 'Ubiquiti', 'Aruba', 'Netgear', 'Seagate',
  'Western Digital', 'Kingston', 'Crucial', 'Anker', 'Xerox', 'Ricoh',
  'Fujitsu', 'Toshiba', 'Zebra', 'Targus', 'Belkin', 'ViewSonic', 'BenQ',
  'Philips', 'Sony', 'Huawei', 'Honor', 'Xiaomi',
];

const MAKE_BY_NAME = new Map(MAKES.map((name) => [name.toUpperCase(), name]));

const SEP = '(?:\\s*[:=]\\s*|\\s+)';

/** Only the labelled forms. An unlabelled name is a guess this file will not make. */
const LABELLED_MAKE = new RegExp(`^(?:MAKE|BRAND|MANUFACTURER)${SEP}(\\S.*)$`, 'i');
const LABELLED_MODEL = new RegExp(`^(?:MODEL(?:\\s*(?:NAME|NO|NUMBER))?|MDL)${SEP}(\\S.*)$`, 'i');

/**
 * What a specification looks like: a quantity with a unit, or the name of
 * a part inside the machine. None of these ever appear in an identifier,
 * which is what makes them safe to route away from the serial number.
 */
const SPEC_PATTERNS = [
  /\b\d+(?:\.\d+)?\s*(?:GB|TB|MB)\b/i,
  /\b(?:RAM|DDR\d?|SSD|HDD|NVMe|eMMC)\b/i,
  /\b(?:Intel|AMD)\b/i,
  /\bCore\s*i[35793]\b/i,
  /\b(?:i[3579]|Ryzen|Celeron|Pentium|Xeon|Athlon|Snapdragon)\b/i,
  /\b\d+(?:\.\d+)?\s*GHz\b/i,
  /\b\d{3,4}\s*x\s*\d{3,4}\b/i,
  /\b(?:FHD|UHD|QHD|WUXGA|IPS|OLED|LED)\b/i,
  /\b\d{2}(?:\.\d)?\s*(?:inch|"|in\b)/i,
  /\b\d+\s*(?:W|mAh|Wh|V)\b/,
  /\b(?:Wi-?Fi|Bluetooth|Ethernet)\b/i,
];

/**
 * A model name reads like words: `ThinkPad T14 Gen 4`. A part number
 * reads like a code. When a line is labelled `Model:` the label says
 * which field it belongs to but not which of those two it is, so the
 * shape decides — and it decides the same way the barcode classifier
 * does, which files a part-shaped `MODEL` code as the part number.
 */
function looksLikeName(value) {
  return partScore(value) < serialScore(value) || /^[A-Za-z][A-Za-z ]{2,}/.test(value);
}

/** One line of recognised text, tidied into something worth reasoning about. */
export function cleanLines(readings) {
  const seen = new Set();
  const kept = [];

  for (const reading of readings ?? []) {
    const text = String(reading?.text ?? '').replace(/\s+/g, ' ').trim();
    if (text.length < MIN_LENGTH) continue;

    // A line with no letter and no digit carries no value, whatever the
    // reader made of the marks it saw.
    if (!/[A-Za-z0-9]/.test(text)) continue;

    // `confidence` is optional: the browser's own text detector does not
    // report one, and a line it found is not to be thrown away for that.
    const confidence = reading?.confidence;
    if (typeof confidence === 'number' && confidence < MIN_CONFIDENCE) continue;

    const key = text.toUpperCase();
    if (seen.has(key)) continue;
    seen.add(key);
    kept.push(text);
  }

  return kept;
}

export function isSpecLine(text) {
  return SPEC_PATTERNS.some((pattern) => pattern.test(text));
}

function readMake(text) {
  const labelled = text.match(LABELLED_MAKE);
  if (labelled) return { value: labelled[1].trim(), guessed: false };

  const known = MAKE_BY_NAME.get(text.toUpperCase());
  // A brand standing alone on its own line. Anything longer is a sentence
  // that happens to contain the name — a warranty note, a web address.
  if (known) return { value: known, guessed: true };

  return null;
}

/**
 * `readings` is `[{ text, confidence }]` as the reader hands them over.
 *
 * Returns the fields it could fill, the lines it could not place, and the
 * names of the fields whose value was worked out rather than read.
 */
export function readTextFields(readings) {
  const cleaned = cleanLines(readings);

  const result = {
    manufacturer: '',
    model: '',
    specSummary: '',
    guessed: [],
  };

  const specs = [];
  const rest = [];

  for (const text of cleaned) {
    const make = readMake(text);
    if (make && !result.manufacturer) {
      result.manufacturer = make.value;
      if (make.guessed) result.guessed.push('manufacturer');
      continue;
    }

    const model = text.match(LABELLED_MODEL);
    if (model) {
      const value = model[1].trim();
      // A labelled model line is claimed either way: as the model when it
      // reads like a name, and otherwise left for `classifyCodes`, which
      // knows the same prefix means the part number.
      if (looksLikeName(value) && !result.model) {
        result.model = value;
        continue;
      }
      rest.push(text);
      continue;
    }

    if (isSpecLine(text)) {
      specs.push(text);
      continue;
    }

    rest.push(text);
  }

  if (specs.length) {
    result.specSummary = specs.join(', ');
    // Nothing on a label announces itself as the specification. Every
    // line here was recognised by its shape, so the whole summary is a
    // guess even when each part of it is obviously right.
    result.guessed.push('specSummary');
  }

  // `format: 'text'` keeps the retail-barcode rule from firing: a line of
  // digits read off a label is not a scanned EAN, and treating it as one
  // would file it as the part number of every identical box.
  const codes = classifyCodes(rest.map((rawValue) => ({ rawValue, format: 'text' })));

  return {
    ...result,
    serialNumber: codes.serialNumber,
    partNumber: codes.partNumber,
    macAddress: codes.macAddress,
    assetTag: codes.assetTag,
    additional: codes.additional,
    guessed: [...result.guessed, ...codes.guessed],
  };
}
