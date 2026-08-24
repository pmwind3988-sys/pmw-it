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
  // `SVC TAG` and `IMEI` name ONE unit, the same as a serial does — a Dell
  // service tag is the serial as far as the register is concerned, and two
  // tablets never share an IMEI.
  { field: 'serialNumber', re: new RegExp(`^(?:S/?N|SER(?:IAL)?(?:\\s*(?:NO|NUM(?:BER)?))?|(?:SVC|SERVICE)\\s*TAG|IMEI(?:\\s*\\d)?)${SEP}(\\S.*)$`, 'i') },
  // `MODEL`, `TYPE` and `EAN` name a MODEL, which is what a part number is
  // for. Every identical tab in the box carries the same one.
  { field: 'partNumber', re: new RegExp(`^(?:P/?N|P/?NO|PART(?:\\s*(?:NO|NUM(?:BER)?))?|(?:MFR\\s*)?MPN|SKU|MODEL|MDL|TYPE|EAN|UPC|ITEM(?:\\s*(?:NO|CODE))?)${SEP}(\\S.*)$`, 'i') },
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
 * An IMEI names one handset or one tablet, never a model — and at fifteen
 * digits it is one character clear of the retail-barcode rule above it, which
 * would otherwise file the only per-unit code on a phone box as a part number.
 */
const IMEI = /^\d{15}$/;

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

/**
 * How much this looks like the MODEL rather than the unit.
 *
 * Scored separately from `serialScore` rather than as its opposite, because
 * the two read different evidence. The punctuation is the strongest of it: a
 * vendor part number carries `#` (HP's `5UF44AA#ABU` locale suffix) or `/`
 * (Apple's `MK2K3LL/A`), and a serial number carries neither — no
 * manufacturer prints a slash in the code that has to be unique per unit.
 *
 * Shortness is evidence too, but weak evidence: plenty of serials are short.
 * It only decides anything when nothing else does.
 */
export function partScore(value) {
  let score = 0;

  if (/[#/]/.test(value)) score += 6;
  // `LC-24B` — a short letter group, a separator, then the variant.
  if (/^[A-Z]{1,5}[-.][A-Z0-9][A-Z0-9.-]*$/i.test(value)) score += 2;
  // Letters with no digits at all is a model name, and no serial scheme in use
  // anywhere would produce it.
  if (!/\d/.test(value)) score += 2;

  if (value.length <= 7) score += 2;
  else if (value.length <= 10) score += 1;
  // Long enough to be unique per unit is long enough not to be a part number.
  if (value.length >= 14) score -= 2;

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
 * `codes` is `[{ rawValue, format, repeat }]` as the detector and the scanning
 * session hand them over. `repeat` means this exact code was already read off
 * a DIFFERENT box in the same session, which is the strongest evidence there
 * is: a serial number appears on one box and one box only, so a code on two of
 * them names the model. That is what separates two tabs bought together — the
 * shared part number is recognised as shared, and each tab keeps its own
 * serial instead of the second one's serial being filed as a part number.
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

    unclaimed.push({
      value: prefixed ? prefixed.value : raw,
      format,
      repeat: entry?.repeat === true,
    });
  }

  // A code seen on an earlier box is the model's, not this unit's. It is taken
  // out before anything is scored, so it can neither become the serial nor
  // lose the part-number slot to a shape heuristic.
  const shared = unclaimed.filter((code) => code.repeat);
  const fresh = unclaimed.filter((code) => !code.repeat);

  if (shared.length && !result.partNumber) {
    result.partNumber = shared.shift().value;
    result.guessed.push('partNumber');
  }

  // Held out of the scoring in both directions: an IMEI must never be filed as
  // a part number, and it must not outrank a serial printed beside it. It
  // takes the serial slot only if nothing else has.
  const imei = fresh.filter((code) => IMEI.test(code.value));
  const shaped = fresh.filter((code) => !IMEI.test(code.value));

  const retail = shaped.filter(isRetail);
  const rest = shaped.filter((code) => !isRetail(code));

  if (retail.length && !result.partNumber) {
    result.partNumber = retail.shift().value;
    result.guessed.push('partNumber');
  }

  // Best first by how much more it looks like a unit than like a model, then
  // length as the tie-break, because between two codes that score alike the
  // longer one carries more to be unique with.
  const lead = (code) => serialScore(code.value) - partScore(code.value);
  rest.sort((a, b) => lead(b) - lead(a) || b.value.length - a.value.length);

  for (const code of rest) {
    // The one code on the box that looks more like a model than a unit is a
    // part number, even with the serial slot standing empty. A box with only
    // `5UF44AA#ABU` printed on it has no serial to read, and inventing one
    // from the part number gives twenty identical items twenty identities.
    const modelShaped = !result.partNumber && lead(code) < 0;

    if (!result.serialNumber && !modelShaped) {
      result.serialNumber = code.value;
      result.guessed.push('serialNumber');
    } else if (!result.partNumber) {
      result.partNumber = code.value;
      result.guessed.push('partNumber');
    } else if (!result.serialNumber) {
      result.serialNumber = code.value;
      result.guessed.push('serialNumber');
    } else {
      result.additional.push(code.value);
    }
  }

  if (imei.length && !result.serialNumber) {
    result.serialNumber = imei.shift().value;
    result.guessed.push('serialNumber');
  }

  // Anything left over is kept verbatim rather than dropped. A code nobody can
  // place today is still the only copy of what was on the box.
  for (const code of imei) result.additional.push(code.value);
  for (const code of shared) result.additional.push(code.value);
  for (const code of retail) result.additional.push(code.value);

  return result;
}
