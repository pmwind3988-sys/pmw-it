/**
 * What each scannable field is called on screen.
 *
 * Beside the scanning logic rather than inside the sheet component,
 * because three screens name these fields — the sheet listing what it
 * read, and the two forms listing what it held back — and a field named
 * two different ways in one flow reads as two different fields.
 */

export const SCAN_FIELD_LABELS = {
  serialNumber: 'Serial number',
  partNumber: 'Part number',
  macAddress: 'MAC address',
  assetTag: 'Asset label',
  manufacturer: 'Make',
  model: 'Model',
  specSummary: 'Specification',
};

export const labelFor = (field) => SCAN_FIELD_LABELS[field] ?? field;
