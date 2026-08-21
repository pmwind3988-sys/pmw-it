const STALE_MS = 180 * 86_400_000;

export function ramBucket(installedRamGB) {
  return typeof installedRamGB === 'number' ? `${installedRamGB} GB` : 'Unknown';
}

export const isStale = (device, now = Date.now()) =>
  typeof device.scannedOn === 'number' && now - device.scannedOn > STALE_MS;

const MATCHERS = {
  risk: (device, value) => device.riskLevel === value,
  type: (device, value) => device.deviceType === value,
  department: (device, value) => (device.department ?? 'Unassigned') === value,
  storage: (device, value) => device.storageType === value,
  ram: (device, value) => ramBucket(device.installedRamGB) === value,
  cpu: (device, value) => device.cpuAgeBand === value,
  windows: (device, value) => device.windowsVersion === value,
  os: (device, value) =>
    (value === 'Unsupported' ? device.osSupported === false : device.osSupported === true),
  av: (device, value) =>
    (value === 'Unprotected' ? !device.avProtected : Boolean(device.avProtected)),
  stale: (device, value) => value !== '1' || isStale(device),
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
