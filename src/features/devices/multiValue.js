/**
 * Splitting the fields that hold more than one thing.
 *
 * The scan reports pack several values into one field, and SharePoint stores
 * them the same way: entries separated by a newline (or a comma, in the case
 * of Office), with `|` separating the parts *within* one entry --
 * `Kingston | DDR4 | 3200 MHz`. So `|` is never an entry separator, and a
 * comma only is where the source used it that way.
 *
 * Only the fields listed here are split. A CPU model ("Intel(R) Core(TM)
 * i5-8250U CPU @ 1.60GHz, 4 cores") contains commas and pipes of its own, and
 * splitting it would invent devices that do not exist.
 */

/** Fields the parser or SharePoint stores as several entries. */
export const MULTI_VALUE_KEYS = new Set([
  'gpuList', 'riskReasons', 'fitReasons', 'microsoftOffice', 'adobeProducts', 'manualFields',
  'antivirusProducts', 'storageDrivesRaw', 'ignoredDrives', 'ramSlotInfoRaw', 'monitorsRaw',
  'serverFolders', 'serverCredentials', 'emailDataFiles', 'extraFields',
]);

/** Fields whose source separates entries with a comma as well as a newline. */
const COMMA_KEYS = new Set(['microsoftOffice', 'adobeProducts']);

export const isMultiValue = (key) => MULTI_VALUE_KEYS.has(key);

const clean = (text) => String(text).trim();

/** `{ product, enabled }` and friends, flattened the way SharePoint stores them. */
const asText = (entry) => {
  if (entry === null || entry === undefined) return '';
  if (typeof entry !== 'object') return clean(entry);
  if ('product' in entry) return `${entry.product} | ${entry.enabled ? 'Enabled' : 'Disabled'}`;
  return Object.values(entry).filter((part) => part !== null && part !== '').join(' | ');
};

/**
 * The entries in a multi-value field, each one already split into its parts.
 * Returns `[]` for anything empty, and for a single-value field returns the
 * one entry -- so a caller can treat every cell the same way.
 */
export function splitEntries(value, key) {
  if (value === null || value === undefined || value === '') return [];

  const raw = Array.isArray(value) ? value.map(asText) : [asText(value)];

  const entries = raw
    .flatMap((text) => text.split('\n'))
    .flatMap((text) => (COMMA_KEYS.has(key) ? text.split(',') : [text]))
    .map(clean)
    .filter(Boolean);

  return entries.map((text) => ({
    text,
    parts: text.split('|').map(clean).filter(Boolean),
  }));
}

/** How many separate things a field holds -- 0, 1, or many. */
export const entryCount = (value, key) => splitEntries(value, key).length;
