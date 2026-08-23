/**
 * Working out which barcode is which.
 *
 * A box has several codes on it and nothing on the label says which is the
 * serial number. Some manufacturers are kind and encode a `S/N:` prefix into
 * the barcode itself; most are not. This scores what is left by shape.
 *
 * Every answer it reaches is marked as guessed and shown as guessed in the
 * review grid, for the same reason the device import flags derived values: a
 * heuristic that cannot be corrected is a heuristic that ships wrong data.
 */

/**
 * A separator is required. Without one, `SNK4820` — a perfectly ordinary
 * serial — reads as the prefix `SN` followed by `K4820`, and a confident wrong
 * answer is worse than no answer.
 */
const SEP = '(?:\\s*[:=#]\\s*|\\s+)';

const PREFIXES = [
  { field: 'serialNumber', re: new RegExp(`^(?:S/?N|SER(?:IAL)?(?:\\s*(?:NO|NUM(?:BER)?))?)${SEP}(\\S.*)$`, 'i') },
  { field: 'partNumber', re: new RegExp(`^(?:P/?N|PART(?:\\s*(?:NO|NUM(?:BER)?))?|MPN|SKU)${SEP}(\\S.*)$`, 'i') },
  { field: 'macAddress', re: new RegExp(`^MAC(?:\\s*ADDR(?:ESS)?)?${SEP}(\\S.*)$`, 'i') },
  { field: 'assetTag', re: new RegExp(`^(?:ASSET(?:\\s*TAG)?|TAG)${SEP}(\\S.*)$`, 'i') },
];

/** Only the separated form. Twelve bare hex characters is also a serial shape. */
const MAC = /^(?:[0-9A-F]{2}[:-]){5}[0-9A-F]{2}$/i;

const RETAIL_FORMATS = new Set(['ean_13', 'ean_8', 'upc_a', 'upc_e']);

/**
 * A retail barcode names the MODEL — every identical monitor on the pallet
 * carries the same one — so it is a part number, never a serial. Treating one
 * as a serial is how twenty monitors become one row.
 */
function isRetail(code) {
  return RETAIL_FORMATS.has(code.format) || /^\d{12,14}$/.test(code.value);
}

/**
 * How much this looks like something that identifies one unit rather than a
 * model: mixed letters and digits, and long enough to be unique.
 */
export function serialScore(value) {
  const hasLetter = /[A-Z]/i.test(value);
  const hasDigit = /\d/.test(value);
  const { length } = value;

  let score = 0;
  if (hasLetter && hasDigit) score += 3;
  if (length >= 8 && length <= 20) score += 2;
  else if (length >= 6 && length <= 24) score += 1;
  if (!hasLetter) score -= 2;
  if (length < 6) score -= 1;

  return score;
}

function readPrefix(raw) {
  for (const { field, re } of PREFIXES) {
    const match = raw.match(re);
    if (match) return { field, value: match[1].trim() };
  }
  return null;
}

/**
 * `codes` is `[{ rawValue, format }]` as the detector hands them over.
 *
 * Returns the fields it could fill, the codes it could not place, and the
 * names of the fields whose value was inferred rather than read.
 */
export function classifyCodes(codes) {
  const result = {
    serialNumber: '',
    partNumber: '',
    macAddress: '',
    assetTag: '',
    additional: [],
    guessed: [],
  };

  const seen = new Set();
  const unclaimed = [];

  for (const entry of codes ?? []) {
    const raw = String(entry?.rawValue ?? '').trim();
    if (!raw || seen.has(raw)) continue;
    seen.add(raw);

    const format = String(entry?.format ?? '').toLowerCase();

    // An explicit prefix is the manufacturer telling us outright. It wins over
    // every rule below and is not a guess.
    const prefixed = readPrefix(raw);
    if (prefixed && !result[prefixed.field]) {
      result[prefixed.field] = prefixed.value;
      continue;
    }

    if (MAC.test(raw) && !result.macAddress) {
      result.macAddress = raw.toUpperCase();
      continue;
    }

    unclaimed.push({ value: prefixed ? prefixed.value : raw, format });
  }

  const retail = unclaimed.filter(isRetail);
  const rest = unclaimed.filter((code) => !isRetail(code));

  if (retail.length && !result.partNumber) {
    result.partNumber = retail.shift().value;
    result.guessed.push('partNumber');
  }

  // Best first: score, then length as the tie-break, because between two codes
  // that score alike the longer one carries more to be unique with.
  rest.sort((a, b) => (
    serialScore(b.value) - serialScore(a.value) || b.value.length - a.value.length
  ));

  for (const code of rest) {
    if (!result.serialNumber) {
      result.serialNumber = code.value;
      result.guessed.push('serialNumber');
    } else if (!result.partNumber) {
      result.partNumber = code.value;
      result.guessed.push('partNumber');
    } else {
      result.additional.push(code.value);
    }
  }

  // Anything left over is kept verbatim rather than dropped. A code nobody can
  // place today is still the only copy of what was on the box.
  for (const code of retail) result.additional.push(code.value);

  return result;
}
