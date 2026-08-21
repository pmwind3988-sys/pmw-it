const STALE_MS = 180 * 86_400_000;

export function ramBucket(installedRamGB) {
  return typeof installedRamGB === 'number' ? `${installedRamGB} GB` : 'Unknown';
}

export const isStale = (device, now = Date.now()) =>
  typeof device.scannedOn === 'number' && now - device.scannedOn > STALE_MS;

/**
 * The one place a missing value gets a name. The charts count by this and the
 * matchers compare against it, so clicking a bar labelled "Unassigned" finds
 * the very rows that were counted into it. A blank string counts as missing:
 * SharePoint hands back "" for a column nobody filled in.
 */
export const UNASSIGNED = 'Unassigned';

export const labelOf = (value) =>
  (value === null || value === undefined || value === '' ? UNASSIGNED : String(value));

/** The risk bands the dashboard's "Need attention" figure is counted from. */
export const ATTENTION_LEVELS = ['Critical', 'High'];

const MATCHERS = {
  risk: (device, value) => device.riskLevel === value,
  attention: (device) => ATTENTION_LEVELS.includes(device.riskLevel),
  type: (device, value) => labelOf(device.deviceType) === value,
  department: (device, value) => labelOf(device.department) === value,
  storage: (device, value) => labelOf(device.storageType) === value,
  ram: (device, value) => ramBucket(device.installedRamGB) === value,
  cpu: (device, value) => labelOf(device.cpuAgeBand) === value,
  windows: (device, value) => labelOf(device.windowsVersion) === value,
  os: (device, value) =>
    (value === 'Unsupported' ? device.osSupported === false : device.osSupported === true),
  // Mirrors the dashboard figure exactly: an antivirus state the scan could not
  // read is unknown, not unprotected, and must not land in either bucket.
  av: (device, value) =>
    (value === 'Unprotected' ? device.avProtected === false : device.avProtected === true),
  stale: (device) => isStale(device),
  q: (device, value) => {
    const needle = value.toLowerCase();
    return `${device.computerName ?? ''} ${device.owner ?? ''}`.toLowerCase().includes(needle);
  },
};

export function applyFilters(devices, params) {
  return devices.filter((device) =>
    Object.entries(params).every(([key, value]) => {
      if (!value) return true;
      const matcher = MATCHERS[key];
      return matcher ? matcher(device, value) : true;
    }));
}

const cell = (value) => {
  if (value === null || value === undefined) return '';
  const text = Array.isArray(value) ? value.join('; ') : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

export function toCsv(devices, columns) {
  const header = columns.map((column) => cell(column.label)).join(',');
  const body = devices.map((device) =>
    columns.map((column) => cell(device[column.key])).join(','));
  return [header, ...body].join('\r\n');
}
