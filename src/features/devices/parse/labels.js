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
  // Added 2026-10-05 so a machine keeps its history through a rename. The scan
  // script writes it as `Serial Number: $((Get-CimInstance Win32_BIOS).SerialNumber)`.
  'Serial Number',
  // Added 2026-10-26 for scan version 2. Each is a block of value lines, one
  // ` | `-separated record per line, and `None` / `None found` means the scan
  // looked and found nothing. A v1 report simply has none of these.
  'Manufacturer',
  'Product Number',
  'Model Family',
  'Motherboard Serial',
  'Device UUID',
  'BIOS',
  'Chassis',
  'Memory Modules',
  'Drive Serials',
  'Monitor Details',
  'Battery',
  'Windows Build',
  'Windows License',
  'Office License',
  'Office Version',
  'Office Accounts',
  // Added 2026-10-26: which account each Microsoft 365 licence belongs to, its
  // tenant, and when the PC next checks it in. Not the plan (Basic/Standard/Premium).
  'Office Subscription',
  'Device Join',
  'Installed Software',
  'Scan Info',
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
