import { DEVICE_COLUMNS } from './sharepoint/deviceSchema';

/**
 * Which columns the device list shows, and the every-column list the CSV
 * export always writes. Pure, so the sets can be tested without a table.
 *
 * The list reads as a person reads a machine: who has it and where first, then
 * how healthy each part is, then the detail behind it. The schema's own order
 * is kept for everything not named here, so a field the scan adds later still
 * surfaces at the end rather than vanishing.
 */

/**
 * Columns worked out on the way in rather than stored. They are not in the
 * SharePoint schema and never will be -- the verdict is recomputed on every
 * read -- so the table has to name them itself.
 */
const CALCULATED_COLUMNS = [
  { key: 'personaLabel', label: 'Workload Profile' },
  { key: 'gradeCpu', label: 'CPU Grade' },
  { key: 'gradeRam', label: 'RAM Grade' },
  { key: 'gradeStorage', label: 'Storage Grade' },
  { key: 'gradeGraphics', label: 'Graphics Grade' },
  { key: 'gradeWindows', label: 'Windows Grade' },
  { key: 'actionRequired', label: 'Action Required' },
  { key: 'suggestedFormFactor', label: 'Suggested Form Factor' },
  { key: 'licenseStatus', label: 'Office Licence' },
  { key: 'licenseNote', label: 'Licence Detail' },
  { key: 'gpuClass', label: 'Graphics' },
  { key: 'networkRisk', label: 'Server Link' },
  { key: 'networkNote', label: 'Server Link Detail' },
];

/** `rawReport` is the whole report file, one cell of which is taller than the screen. */
const SKIP_KEYS = new Set(['scannedOn', 'scannedOnMYT', 'rawReport']);

/** Labels the table says differently from the schema. */
const LABEL_OVERRIDES = { riskLevel: 'Security risk' };

/** The columns worth seeing first, in this order. */
const LEAD_KEYS = [
  'computerName', 'status', 'owner', 'location', 'department', 'personaLabel', 'gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows', 'actionRequired',
  'suggestedFormFactor', 'deviceType', 'licenseStatus', 'computerModel', 'cpuModel',
  'installedRamGB', 'storageTotalGB', 'storageType', 'windowsVersion',
  'antivirusStatus', 'riskScore', 'riskLevel',
];

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

const rank = (key) => {
  const lead = LEAD_KEYS.indexOf(key);
  return lead === -1 ? LEAD_KEYS.length : lead;
};

/** Every column the device has, in display order. Scanned is last, as it always was. */
export const ALL_COLUMNS = [
  // `Title` holds the computer name, so the schema never lists it.
  { key: 'computerName', label: 'Computer' },
  ...CALCULATED_COLUMNS,
  ...DEVICE_COLUMNS
    .map((column) => ({
      key: keyFor(column.StaticName),
      label: column.Title,
      kind: column.kind,
      numeric: column.kind === 'number',
    }))
    .filter((column) => !SKIP_KEYS.has(column.key)),
  { key: 'scannedOn', label: 'Scanned', kind: 'datetime' },
]
  .map((column) => (LABEL_OVERRIDES[column.key] ? { ...column, label: LABEL_OVERRIDES[column.key] } : column))
  // A stable sort, so the columns outside LEAD_KEYS keep their schema order.
  .sort((a, b) => rank(a.key) - rank(b.key));

/** The identity every set shares. The computer name is always drawn in the frozen cell. */
const IDENTITY = ['status', 'owner', 'location', 'department'];
const GRADES = ['gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows'];

export const DEFAULT_SET = 'essentials';

/**
 * The four column sets, in the order the picker offers them. `keys` are the
 * columns BESIDES the frozen computer name, in the order they are drawn.
 */
export const COLUMN_SETS = {
  essentials: {
    label: 'Essentials',
    keys: [...IDENTITY, ...GRADES, 'scannedOn'],
  },
  hardware: {
    label: 'Hardware',
    keys: [
      ...IDENTITY,
      'deviceType', 'computerModel', 'cpuModel', 'cpuGeneration', 'installedRamGB',
      'ramSlotsUsed', 'storageTotalGB', 'storageType', 'gpuClass', 'scannedOn',
    ],
  },
  security: {
    label: 'Security',
    keys: [
      ...IDENTITY,
      'windowsVersion', 'osSupported', 'antivirusStatus', 'licenseStatus', 'riskLevel',
      'riskScore', 'networkRisk', 'scannedOn',
    ],
  },
  everything: {
    label: `Everything (${ALL_COLUMNS.length - 1})`,
    keys: ALL_COLUMNS.map((column) => column.key).filter((key) => key !== 'computerName'),
  },
};

export const COLUMN_SET_IDS = Object.keys(COLUMN_SETS);

/** The column objects a set draws, in its own order. Unknown keys are dropped rather than drawn blank. */
export function columnsForSet(setId, all = ALL_COLUMNS) {
  const set = COLUMN_SETS[setId] ?? COLUMN_SETS[DEFAULT_SET];
  const byKey = new Map(all.map((column) => [column.key, column]));
  return set.keys.map((key) => byKey.get(key)).filter(Boolean);
}

/** Only a value this module knows is accepted back from storage. */
export function parseColumnSet(stored) {
  return COLUMN_SET_IDS.includes(stored) ? stored : DEFAULT_SET;
}
