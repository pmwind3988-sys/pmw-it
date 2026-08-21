import { cleanValue } from '../parse/placeholders.js';
import { parseAntivirus } from '../parse/parseValues.js';

const INACTIVE = 'Installed — Inactive';

function readWindows(lines) {
  const windowsVersion = lines.length ? cleanValue(lines[0]) : null;
  if (!windowsVersion) {
    return { windowsVersion: null, windowsMajor: null, windowsEdition: null, osSupported: null };
  }

  const match = /Windows\s+(\d+)\s*(.*)$/i.exec(windowsVersion);
  const windowsMajor = match ? Number(match[1]) : null;

  return {
    windowsVersion,
    windowsMajor,
    windowsEdition: match && match[2] ? match[2].trim() : null,
    // Windows 10 reached end of support on 14 October 2025.
    osSupported: windowsMajor === null ? null : windowsMajor >= 11,
  };
}

function readAntivirusStatus(raw, products) {
  if (raw) {
    // `DEACTIVATED` contains `ACTIVAT`, so it has to be tested before `activ`.
    if (/not\s*installed/i.test(raw)) return 'Not Installed';
    if (/deactivat|disabled|expired/i.test(raw)) return INACTIVE;
    if (/\d+\s*days?|trial/i.test(raw)) return 'Trial';
    if (/activ|enabled/i.test(raw)) return 'Active';
  }

  if (!products.length) return 'Unknown';
  return products.some((entry) => entry.enabled) ? 'Active' : INACTIVE;
}

export function deriveHealth(fields) {
  const antivirusProducts = parseAntivirus(fields.Antivirus ?? []);
  const antivirusStatusRaw = fields['Antivirus status']?.length
    ? cleanValue(fields['Antivirus status'][0])
    : null;

  // A scan that failed early writes the header and nothing else. Importing it
  // as a machine with no CPU and no disk would drag down every fleet average,
  // so it is marked instead and excluded from the statistics.
  const scanComplete = !(
    !fields['Computer Name']?.length
    && !fields.Processor?.length
    && !fields['Storage Drives']?.length
  );

  return {
    ...readWindows(fields['Windows Version'] ?? []),
    antivirusStatus: readAntivirusStatus(antivirusStatusRaw, antivirusProducts),
    antivirusStatusRaw,
    antivirusProducts,
    avProtected: antivirusProducts.some((entry) => entry.enabled),
    scanComplete,
  };
}
