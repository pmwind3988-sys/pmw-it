import { DEVICE_COLUMNS } from './sharepoint/deviceSchema.js';

/**
 * How one device's fields are laid out on its own page: by what somebody is
 * looking for -- the specs, the mailboxes, what it can reach on the server --
 * rather than by the order SharePoint happens to store them in.
 *
 * Every field the schema produces belongs to exactly one group. Anything added
 * to the schema and not listed here lands in "Other", so a new column is never
 * silently invisible; `fieldGroups.test.js` is what notices.
 */
const GROUPS = [
  {
    id: 'identity',
    title: 'Identity',
    hint: 'Who has it and where it lives',
    keys: ['computerName', 'owner', 'ownerSource', 'department', 'deviceType', 'anydeskId'],
  },
  {
    id: 'risk',
    title: 'Risk',
    hint: 'Why this machine does or does not need attention',
    keys: ['riskScore', 'riskLevel', 'riskReasons', 'scanComplete', 'remarks'],
  },
  {
    id: 'specs',
    title: 'Specs',
    hint: 'The machine itself',
    keys: [
      'computerModel', 'motherboardVendor', 'motherboardModel',
      'cpuModel', 'cpuVendor', 'cpuGeneration', 'cpuArchitecture',
      'cpuGenerationRank', 'cpuAgeBand',
    ],
  },
  {
    id: 'memory',
    title: 'Memory',
    keys: [
      'installedRamGB', 'reportedRamGB', 'ramDiscrepancy', 'ramType', 'ramSpeedMhz',
      'ramSlotsUsed', 'ramSlotsTotal', 'ramUpgradable', 'ramSlotInfoRaw',
    ],
  },
  {
    id: 'storage',
    title: 'Storage',
    keys: [
      'storageTotalGB', 'driveCount', 'storageType', 'hasHdd',
      'storageDrivesRaw', 'ignoredDrives',
    ],
  },
  {
    id: 'os',
    title: 'Operating system',
    keys: ['windowsVersion', 'windowsMajor', 'windowsEdition', 'osSupported'],
  },
  {
    id: 'security',
    title: 'Security',
    keys: ['antivirusStatus', 'antivirusStatusRaw', 'antivirusProducts', 'avProtected'],
  },
  {
    id: 'network',
    title: 'Network',
    keys: ['networkType', 'ssid', 'ipAddress', 'ipAssignment'],
  },
  {
    id: 'display',
    title: 'Display',
    keys: ['gpuList', 'monitorCount', 'monitorsRaw'],
  },
  {
    id: 'software',
    title: 'Software',
    keys: ['microsoftOffice', 'adobeProducts'],
  },
  {
    id: 'email',
    title: 'Email',
    hint: 'Mailboxes and the data files behind them',
    keys: ['mailboxCount', 'archiveCount', 'emailDataFiles'],
  },
  {
    id: 'server',
    title: 'Server access',
    keys: ['mappedDrives', 'serverFolders', 'serverCredentials'],
  },
  {
    id: 'record',
    title: 'Record',
    hint: 'Where this row came from',
    keys: [
      'scannedOn', 'scannedOnMYT', 'importedOn', 'sourceFileName',
      'manualFields', 'extraFields',
    ],
  },
];

/** The whole report file. It has its own panel, not a row in a group. */
export const RAW_REPORT_KEY = 'rawReport';

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

const LABELS = new Map([
  ['computerName', 'Computer'],
  ...DEVICE_COLUMNS.map((column) => [keyFor(column.StaticName), column.Title]),
]);

const KINDS = new Map(DEVICE_COLUMNS.map((column) => [keyFor(column.StaticName), column.kind]));

export const labelFor = (key) => LABELS.get(key) ?? key;

/** Every key the schema knows, minus the raw report. */
export const ALL_KEYS = [
  'computerName',
  ...DEVICE_COLUMNS.map((column) => keyFor(column.StaticName)),
].filter((key) => key !== RAW_REPORT_KEY);

const GROUPED = new Set(GROUPS.flatMap((group) => group.keys));

const other = ALL_KEYS.filter((key) => !GROUPED.has(key));

/**
 * The groups, each field carrying its label and kind, and each group knowing
 * whether this device has anything to show in it -- an empty group is a row of
 * dashes nobody needs.
 */
export const FIELD_GROUPS = [
  ...GROUPS,
  ...(other.length ? [{ id: 'other', title: 'Other', keys: other }] : []),
].map((group) => ({
  ...group,
  fields: group.keys.map((key) => ({
    key,
    label: labelFor(key),
    kind: KINDS.get(key) ?? 'text',
  })),
}));

const filled = (value) => value !== null && value !== undefined && value !== '';

/** The groups worth drawing for one device, in order. */
export function groupsFor(device, { includeEmpty = false } = {}) {
  return FIELD_GROUPS
    .map((group) => ({
      ...group,
      fields: includeEmpty ? group.fields : group.fields.filter((f) => filled(device?.[f.key])),
    }))
    .filter((group) => group.fields.length > 0);
}
