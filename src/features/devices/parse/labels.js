/**
 * The labels the scan script writes, in file order. Verified identical across
 * all 17 sample reports.
 *
 * This list is the ONLY thing allowed to open a new field. A generic /^\w+:/
 * split reads "Total Slots: 2 | Used Slots: 2" and "Y: | \\server\PMW\IT" as
 * fields, which silently moves those values out of the block they belong to.
 */
export const KNOWN_LABELS = [
  'Name',
  'Anydesk',
  'Antivirus status',
  'Remarks',
  'Computer Name',
  'Computer Model',
  'Motherboard',
  'Windows Version',
  'Processor',
  'GPU',
  'Total RAM',
  'RAM Slot Info',
  'Storage Drives',
  'Network Information',
  'Antivirus',
  'Monitor',
  'PMW Server and credentials',
  'Server folder',
  'Microsoft Office',
  'Adobe',
  'Email data files found Active or Inactive account',
];

const normalise = (s) => s.replace(/\s+/g, ' ').trim().toLowerCase();

const BY_NORMALISED = new Map(KNOWN_LABELS.map((label) => [normalise(label), label]));

/**
 * Returns the canonical label and any inline value, or null when the line is
 * not a label. Splits on the FIRST colon only, so an inline value may itself
 * contain colons.
 */
export function matchLabel(line) {
  const colon = line.indexOf(':');
  if (colon === -1) return null;

  const label = BY_NORMALISED.get(normalise(line.slice(0, colon)));
  if (!label) return null;

  return { label, inline: line.slice(colon + 1).trim() };
}
