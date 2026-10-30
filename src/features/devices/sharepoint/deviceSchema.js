import { formatMYT } from '../../../utils/malaysiaTime.js';
import { STATUSES } from '../lifecycle/status.js';
import { WINDOWS_LICENSE_TYPES, OFFICE_LICENSE_TYPES, DEVICE_JOIN_STATES } from '../derive/deriveLicensing.js';

export const DEVICE_LIST_NAME = 'IT Device List';
export const CHANGE_LIST_NAME = 'IT Device Changes';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
/**
 * `key` names the record field when it is not the StaticName's camelCase. A
 * raw text column whose parsed form is an array keeps the raw text under its
 * own key, so the array and the text can both sit on a record.
 */
const note = (StaticName, Title, key) => ({ StaticName, Title, kind: 'note', ...(key ? { key } : {}) });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const bool = (StaticName, Title) => ({ StaticName, Title, kind: 'boolean' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

const RISK_LEVELS = ['Critical', 'High', 'Watch', 'OK', 'Unknown'];

/** `Title` holds the computer name and is built in, so it is never created. */
export const DEVICE_COLUMNS = [
  text('Owner', 'Owner'),
  choice('OwnerSource', 'Owner Source',
    ['Name field', 'Filename', 'Server credential', 'Email', 'Manual']),
  text('Department', 'Department'),
  text('Location', 'Location'),
  choice('DeviceType', 'Device Type', ['Laptop', 'Desktop', 'Unknown']),
  text('ComputerModel', 'Model'),
  text('MotherboardVendor', 'Motherboard Vendor'),
  text('MotherboardModel', 'Motherboard Model'),
  text('AnydeskId', 'AnyDesk ID'),
  text('SerialNumber', 'Serial Number'),
  choice('Status', 'Status', STATUSES),
  date('StatusChangedOn', 'Status Changed On'),

  date('ScannedOn', 'Scanned On'),
  date('ImportedOn', 'Imported On'),
  text('ScannedOnMYT', 'Scanned On (MYT)'),
  text('SourceFileName', 'Source File'),

  text('WindowsVersion', 'Windows Version'),
  num('WindowsMajor', 'Windows Major'),
  text('WindowsEdition', 'Windows Edition'),
  bool('OsSupported', 'OS Supported'),

  text('CpuModel', 'CPU'),
  choice('CpuVendor', 'CPU Vendor', ['Intel', 'AMD', 'Other']),
  text('CpuGeneration', 'CPU Generation'),
  text('CpuArchitecture', 'CPU Architecture'),
  num('CpuGenerationRank', 'CPU Generation (Intel-equivalent)'),
  choice('CpuAgeBand', 'CPU Age', ['Current', 'Aging', 'Obsolete', 'Unknown']),

  num('InstalledRamGB', 'Installed RAM (GB)'),
  num('ReportedRamGB', 'Reported RAM (GB)'),
  bool('RamDiscrepancy', 'RAM Discrepancy'),
  text('RamType', 'RAM Type'),
  num('RamSpeedMhz', 'RAM Speed (MHz)'),
  num('RamSlotsUsed', 'RAM Slots Used'),
  num('RamSlotsTotal', 'RAM Slots Total'),
  bool('RamUpgradable', 'RAM Upgradable'),
  note('RamSlotInfoRaw', 'RAM Slot Info'),

  num('StorageTotalGB', 'Storage Total (GB)'),
  num('DriveCount', 'Drive Count'),
  choice('StorageType', 'Storage Type', ['SSD only', 'Mixed', 'HDD only', 'Unknown']),
  bool('HasHdd', 'Has HDD'),
  note('StorageDrivesRaw', 'Storage Drives'),
  note('IgnoredDrives', 'Drives Not Counted'),

  choice('AntivirusStatus', 'Antivirus Status',
    ['Active', 'Installed — Inactive', 'Trial', 'Not Installed', 'Unknown']),
  text('AntivirusStatusRaw', 'Antivirus Status (raw)'),
  note('AntivirusProducts', 'Antivirus Products'),
  bool('AvProtected', 'Protected'),

  text('NetworkType', 'Network'),
  text('Ssid', 'SSID'),
  text('IpAddress', 'IP Address'),
  choice('IpAssignment', 'IP Assignment', ['Dynamic', 'Static', 'Unknown']),

  note('GpuList', 'GPU'),
  num('MonitorCount', 'Monitors'),
  note('MonitorsRaw', 'Monitors (raw)'),

  note('MicrosoftOffice', 'Microsoft Office'),
  note('AdobeProducts', 'Adobe'),
  num('MappedDrives', 'Mapped Drives'),
  note('ServerFolders', 'Server Folders'),
  note('ServerCredentials', 'Server Credentials'),

  num('MailboxCount', 'Mailboxes'),
  num('ArchiveCount', 'Archives'),
  note('EmailDataFiles', 'Email Data Files'),

  // Scan report version 2. A v1 report leaves all of these empty, and a
  // re-import of a v1 report never clears them: see carryV2Fields in syncDevices.js.
  text('Manufacturer', 'Manufacturer'),
  text('ProductNumber', 'Product Number'),
  text('ModelFamily', 'Model Family'),
  text('MotherboardSerial', 'Motherboard Serial'),
  text('DeviceUuid', 'Device UUID'),
  text('BiosVersion', 'BIOS Version'),
  text('BiosDate', 'BIOS Date'),
  text('Chassis', 'Chassis'),
  text('AssetTag', 'Asset Tag'),
  note('MemoryModules', 'Memory Modules', 'memoryModulesRaw'),
  note('DriveSerials', 'Drive Serials', 'driveSerialsRaw'),
  note('MonitorDetails', 'Monitor Details', 'monitorDetailsRaw'),
  text('WindowsBuild', 'Windows Build'),

  choice('WindowsLicenseType', 'Windows Licence', WINDOWS_LICENSE_TYPES),
  choice('WindowsActivation', 'Windows Activation', ['Activated', 'Not activated', 'Unknown']),
  text('WindowsKeyEnds', 'Windows Key Ends'),
  text('FirmwareKey', 'Firmware Key'),
  choice('OfficeLicenseType', 'Office', OFFICE_LICENSE_TYPES),
  bool('OfficePersonal', 'Office Personal'),
  note('OfficePersonalReason', 'Office Personal Reason'),
  note('OfficeSignedIn', 'Office Signed-in Licences', 'officeSignedInLicences'),
  note('OfficeAccounts', 'Office Accounts', 'officeAccountsRaw'),
  text('OfficeLicensedTo', 'Office Licensed To'),
  text('OfficeTenantId', 'Office Tenant ID'),
  text('OfficeLicenceValidUntil', 'Office Licence Check-in By'),
  note('OfficeSubscriptions', 'Office Subscriptions', 'officeSubscriptionsRaw'),
  text('OfficeVersion', 'Office Version'),
  text('OfficeChannel', 'Office Channel'),
  choice('DeviceJoin', 'Joined To', DEVICE_JOIN_STATES),
  text('TenantName', 'Tenant'),

  note('InstalledSoftware', 'Installed Software', 'installedSoftwareRaw'),
  num('SoftwareCount', 'Programs Installed'),
  // Text, not a choice: a scan version this code has not met yet must not be refused by SharePoint.
  text('ScanVersion', 'Scan Version'),
  bool('ScanRunAsAdmin', 'Scan Run As Admin'),

  num('RiskScore', 'Risk Score'),
  choice('RiskLevel', 'Risk Level', RISK_LEVELS),
  note('RiskReasons', 'Risk Reasons'),
  bool('ScanComplete', 'Scan Complete'),
  note('Remarks', 'Remarks'),
  note('ManualFields', 'Manually Set Fields'),
  // Names it answered to before a rename, oldest first. The change log holds
  // the same as "Renamed" lines; this copy is what the List and search read.
  note('PreviousNames', 'Previous Names'),
  note('ExtraFields', 'Extra Fields'),
  note('RawReport', 'Raw Report'),
];

export const CHANGE_COLUMNS = [
  num('DeviceId', 'Device ID'),
  text('FieldName', 'Field'),
  note('OldValue', 'Old Value'),
  note('NewValue', 'New Value'),
  date('ChangedOn', 'Changed On'),
  text('ChangedOnMYT', 'Changed On (MYT)'),
  text('ChangedBy', 'Changed By'),
  choice('ChangeType', 'Change Type', ['Added', 'Updated', 'Removed']),
];

/**
 * Only these produce change-log rows. IP address, SSID and mapped drives are
 * deliberately absent: they are DHCP-assigned or session-dependent and change
 * constantly, and logging them would bury the hardware changes that matter.
 *
 * `manualFields` is absent for a different reason: it records how a value was
 * set, not that the machine changed.
 *
 * computerName, serialNumber and status joined when machines got an identity beyond their name: a rename, a first serial and a move to the stash are exactly what the history is for.
 */
export const TRACKED_FIELDS = [
  'computerName', 'location', 'serialNumber', 'status',
  'owner', 'department', 'deviceType', 'computerModel',
  'windowsVersion', 'osSupported',
  'cpuModel', 'cpuAgeBand',
  'installedRamGB', 'ramType', 'ramSlotsUsed',
  'storageTotalGB', 'storageType',
  'antivirusStatus', 'riskLevel',
  // Version 2. Only the ones that say something about the machine's state.
  // The memory and drive lines change only when a stick or drive is swapped
  // (or, rarely, a drive's firmware is updated), which is exactly what the
  // machine page promises Spec history will show. The Windows build number and
  // the software list change with every update and would bury the log.
  'productNumber', 'windowsLicenseType', 'windowsActivation',
  'officeLicenseType', 'officePersonal', 'deviceJoin', 'softwareCount',
  'memoryModulesRaw', 'driveSerialsRaw',
];

/** The columns a scan version 2 report fills, by StaticName. */
const V2_STATIC_NAMES = [
  'Manufacturer', 'ProductNumber', 'ModelFamily', 'MotherboardSerial', 'DeviceUuid',
  'BiosVersion', 'BiosDate', 'Chassis', 'AssetTag', 'MemoryModules', 'DriveSerials',
  'MonitorDetails', 'WindowsBuild', 'WindowsLicenseType', 'WindowsActivation',
  'WindowsKeyEnds', 'FirmwareKey', 'OfficeLicenseType', 'OfficePersonal',
  'OfficePersonalReason', 'OfficeSignedIn', 'OfficeAccounts', 'OfficeLicensedTo',
  'OfficeTenantId', 'OfficeLicenceValidUntil', 'OfficeSubscriptions', 'OfficeVersion',
  'OfficeChannel', 'DeviceJoin', 'TenantName', 'InstalledSoftware', 'SoftwareCount',
  'ScanVersion', 'ScanRunAsAdmin',
];

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

/** The record field a column reads and writes. */
export const recordKey = (column) => column.key ?? keyFor(column.StaticName);

/** Record fields that only a version 2 scan fills. Carried over on a v1 re-import. */
export const V2_FIELDS = DEVICE_COLUMNS
  .filter((column) => V2_STATIC_NAMES.includes(column.StaticName))
  .map(recordKey);

const serialise = (value) => {
  if (Array.isArray(value)) {
    return value
      .map((entry) =>
        (entry && typeof entry === 'object' && 'product' in entry
          ? `${entry.product} | ${entry.enabled ? 'Enabled' : 'Disabled'}`
          : String(entry)))
      .join('\n');
  }
  return value == null ? '' : String(value);
};

export function toListItem(device) {
  const item = { Title: device.computerName ?? '' };

  for (const column of DEVICE_COLUMNS) {
    let value = device[recordKey(column)];

    if (column.StaticName === 'ScannedOnMYT') {
      value = device.scannedOn ? formatMYT(device.scannedOn, 'datetime12') : null;
    }
    if (column.StaticName === 'ExtraFields') {
      value = device.unknownLabels?.length ? JSON.stringify(device.unknownLabels) : null;
    }

    switch (column.kind) {
      case 'text':
      case 'note':
        // Empty string clears the column; null would be rejected.
        item[column.StaticName] = serialise(value);
        break;
      case 'number':
        if (typeof value === 'number' && Number.isFinite(value)) item[column.StaticName] = value;
        break;
      case 'boolean':
        // `false` is a real value and must survive; only null/undefined is absent.
        if (typeof value === 'boolean') item[column.StaticName] = value;
        break;
      case 'choice':
        if (value) item[column.StaticName] = String(value);
        break;
      case 'datetime':
        if (typeof value === 'number' && Number.isFinite(value)) {
          item[column.StaticName] = new Date(value).toISOString();
        }
        break;
      default:
        break;
    }
  }

  return item;
}

const ARRAY_COLUMNS = new Set([
  'GpuList', 'RiskReasons', 'MicrosoftOffice', 'AdobeProducts', 'ManualFields',
  'OfficeSignedIn', 'PreviousNames',
]);

export function fromListItem(row) {
  const record = { id: row.Id ?? row.ID ?? null, computerName: row.Title ?? null };

  for (const column of DEVICE_COLUMNS) {
    const key = recordKey(column);
    const raw = row[column.StaticName];

    // An absent column reads as null for every kind — notably NOT as NaN for a
    // date, which is what `new Date(undefined).getTime()` would produce.
    if (raw === undefined || raw === null || raw === '') {
      record[key] = null;
      continue;
    }

    if (column.kind === 'datetime') record[key] = new Date(raw).getTime();
    else if (ARRAY_COLUMNS.has(column.StaticName)) record[key] = String(raw).split('\n');
    else record[key] = raw;
  }

  // SharePoint's own creation stamp: the earliest this register knew of the
  // machine, which is what "since at least" means on a pre-history owner.
  record.createdOn = row.Created ? new Date(row.Created).getTime() : null;

  return record;
}
