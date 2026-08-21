import { formatMYT } from '../../datastudio/time/malaysiaTime.js';

export const DEVICE_LIST_NAME = 'IT Device List';
export const CHANGE_LIST_NAME = 'IT Device Changes';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
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
  choice('DeviceType', 'Device Type', ['Laptop', 'Desktop', 'Unknown']),
  text('ComputerModel', 'Model'),
  text('MotherboardVendor', 'Motherboard Vendor'),
  text('MotherboardModel', 'Motherboard Model'),
  text('AnydeskId', 'AnyDesk ID'),

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

  num('RiskScore', 'Risk Score'),
  choice('RiskLevel', 'Risk Level', RISK_LEVELS),
  note('RiskReasons', 'Risk Reasons'),
  bool('ScanComplete', 'Scan Complete'),
  note('Remarks', 'Remarks'),
  note('ExtraFields', 'Extra Fields'),
  note('RawReport', 'Raw Report'),
];

export const CHANGE_COLUMNS = [
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
 */
export const TRACKED_FIELDS = [
  'owner', 'department', 'deviceType', 'computerModel',
  'windowsVersion', 'osSupported',
  'cpuModel', 'cpuAgeBand',
  'installedRamGB', 'ramType', 'ramSlotsUsed',
  'storageTotalGB', 'storageType',
  'antivirusStatus', 'riskLevel',
];

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

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
    const key = keyFor(column.StaticName);
    let value = device[key];

    if (column.StaticName === 'ScannedOnMYT') value = formatMYT(device.scannedOn, 'datetime12');
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

const ARRAY_COLUMNS = new Set(['GpuList', 'RiskReasons', 'MicrosoftOffice', 'AdobeProducts']);

export function fromListItem(row) {
  const record = { id: row.Id ?? row.ID ?? null, computerName: row.Title ?? null };

  for (const column of DEVICE_COLUMNS) {
    const key = keyFor(column.StaticName);
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

  return record;
}
