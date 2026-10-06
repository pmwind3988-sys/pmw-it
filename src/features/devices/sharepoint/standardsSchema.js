/**
 * Every saved standard is a new row; nothing is ever updated or deleted, so the
 * list IS the history and "restore" is just saving an old one again.
 */
export const STANDARDS_LIST_NAME = 'IT Device Standards';

export const STANDARDS_COLUMNS = [
  { StaticName: 'Standard', Title: 'Standard', kind: 'note' },
  { StaticName: 'SavedOn', Title: 'Saved On', kind: 'datetime' },
  { StaticName: 'SavedBy', Title: 'Saved By', kind: 'text' },
  { StaticName: 'Summary', Title: 'Summary', kind: 'note' },
];

export function toStandardItem({ version, standard, savedBy, summary, savedOn }) {
  return {
    Title: `v${version}`,
    Standard: JSON.stringify(standard),
    SavedOn: new Date(savedOn).toISOString(),
    SavedBy: savedBy ?? '',
    Summary: summary ?? '',
  };
}

export function fromStandardItem(row) {
  let standard;
  try {
    standard = JSON.parse(row.Standard);
  } catch {
    standard = null;
  }
  const version = Number(String(row.Title ?? '').replace(/^v/, ''));
  return {
    id: row.Id ?? row.ID ?? null,
    version: Number.isFinite(version) && version > 0 ? version : null,
    standard,
    savedBy: row.SavedBy || null,
    savedOn: row.SavedOn ? new Date(row.SavedOn).getTime() : null,
    summary: row.Summary || null,
  };
}
