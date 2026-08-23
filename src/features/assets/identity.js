import { TRACKED } from './assetKinds.js';

/**
 * What makes two rows the same asset.
 *
 * Every asset carries an `AssetKey`, and the key is what a save upserts on. It
 * is why scanning the same laptop twice updates one row instead of making two,
 * and why a second bag of the same mice adds to a quantity instead of starting
 * a rival line.
 *
 * The key is derived, never typed. A user who retypes a serial number changes
 * the identity of the row, which is correct: it was the wrong serial.
 */

/**
 * Codes — serials, part numbers, sticker labels. Case and spacing on a printed
 * label are not information: `cn0abc 123` and `CN0ABC123` are the same serial
 * read by two different scanners.
 */
export function normaliseCode(value) {
  return String(value ?? '').replace(/\s+/g, '').toUpperCase();
}

/**
 * Names — manufacturer, model, category. Inner spaces are kept because they
 * separate words ("ThinkPad T14 Gen 3"), but runs of them are not meaningful.
 */
export function normaliseName(value) {
  return String(value ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
}

/** `local:` keys are per-row and can never match anything else. */
export const LOCAL_PREFIX = 'local:';

/**
 * A key that identifies the same physical thing next time, or a `local:` one
 * that admits it cannot.
 *
 * Tracked items are identified by manufacturer and serial. Failing a serial
 * they fall back to the sticker label, which is at least unique by policy
 * (§4.7). Failing both there is nothing durable to key on, so the row gets its
 * own local id — deliberately visible as such, because the alternative is a
 * key that silently collides with every other unlabelled unserialised laptop.
 *
 * Bulk lines are identified by what they are, not which one: category,
 * manufacturer and model.
 */
export function assetKey(draft) {
  const manufacturer = normaliseName(draft?.manufacturer);

  if (draft?.trackingMode === TRACKED) {
    const serial = normaliseCode(draft?.serialNumber);
    if (serial) return `serial:${manufacturer}|${serial}`;

    const tag = normaliseCode(draft?.assetTag);
    if (tag) return `tag:${tag}`;

    return `${LOCAL_PREFIX}${draft?.localId ?? ''}`;
  }

  const category = normaliseName(draft?.category);
  const model = normaliseName(draft?.model);
  return `bulk:${category}|${manufacturer}|${model}`;
}

/**
 * Whether re-scanning this thing next month would find the same row. A local
 * key would not, and the review grid says so rather than letting a batch of
 * anonymous rows pile up unnoticed.
 */
export function hasStableIdentity(key) {
  return Boolean(key) && !key.startsWith(LOCAL_PREFIX);
}

/** The register indexed the way a save looks things up. */
export function indexByKey(assets) {
  const index = new Map();
  for (const asset of assets) {
    if (asset?.assetKey) index.set(asset.assetKey, asset);
  }
  return index;
}

/**
 * Sticker labels indexed for the uniqueness check. Normalised on both sides,
 * so a label typed with a stray space still collides with the one on the wall.
 */
export function indexByTag(assets) {
  const index = new Map();
  for (const asset of assets) {
    const tag = normaliseCode(asset?.assetTag);
    if (tag) index.set(tag, asset);
  }
  return index;
}

/**
 * What the row is called in SharePoint. Readable, never load-bearing — a name
 * that changes when somebody corrects a model number must not move the row.
 */
export function assetTitle(draft) {
  const name = [draft?.manufacturer, draft?.model]
    .map((part) => String(part ?? '').trim())
    .filter(Boolean)
    .join(' ');

  const serial = String(draft?.serialNumber ?? '').trim();
  if (name && serial) return `${name} — ${serial}`;
  if (name) return name;
  if (serial) return serial;
  return String(draft?.category ?? 'Unidentified item');
}
