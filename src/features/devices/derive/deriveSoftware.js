import { cleanValue } from '../parse/placeholders.js';
import { splitParts, linesOf, meaningful } from '../parse/parts.js';

/**
 * Everything in Programs and Features, one line per program: name, version,
 * publisher, and the day it was installed when Windows recorded one.
 *
 * Many of these are not choices anybody made: a runtime that a game or an app
 * installed for itself, a driver package, the Windows SDK pieces. They are
 * real and stay in the list, but `isBackgroundComponent` lets the page put them
 * out of the way so the software a person chose is what they see first.
 */

const DATE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Runtimes, redistributables, SDK bits and driver packages. Matched on the name,
 * and on the publisher for the chipset and driver vendors. Developer SDKs such
 * as `Microsoft .NET SDK` are NOT here: a developer chose to install one.
 */
const BACKGROUND_NAMES = [
  /Visual C\+\+.*Redistributable/i,
  /\.NET Runtime\b/i,
  /Desktop Runtime\b/i,
  /Shared Framework\b/i,
  /\.NET Host\b/i,
  /\bWindows (SDK|Software Development Kit)\b/i,
  /Update Health Tools/i,
  /WebView2 Runtime/i,
  /Application Compatibility Fix Database/i,
];

const VENDOR_PUBLISHER = /^(Intel|Realtek|AMD|NVIDIA)\b/i;
const VENDOR_DRIVER_NAME = /driver|install|chipset|graphics|audio|management engine|software package/i;

export function isBackgroundComponent(name, publisher = '') {
  const text = String(name ?? '');
  if (BACKGROUND_NAMES.some((pattern) => pattern.test(text))) return true;
  return VENDOR_PUBLISHER.test(String(publisher ?? '').trim())
    && VENDOR_DRIVER_NAME.test(text);
}

/** Yyyy-mm-dd, or null when the scan left the date blank. */
const day = (cell) => (DATE.test(cell ?? '') ? cell : null);

export function parseSoftware(lines) {
  return meaningful(lines)
    .map((line) => {
      const [name, version, publisher, installed] = splitParts(line);
      return {
        name: cleanValue(name),
        version: cleanValue(version),
        publisher: cleanValue(publisher),
        installedOn: day(installed),
      };
    })
    .filter((entry) => entry.name);
}

export function deriveSoftware(fields) {
  const block = fields['Installed Software'] ?? [];
  const installedSoftware = parseSoftware(block);
  const kept = meaningful(block);

  return {
    installedSoftware,
    installedSoftwareRaw: kept.length ? kept.join('\n') : null,
    // Null for an old report: the list was never taken, so zero would be a lie.
    softwareCount: block.length ? installedSoftware.length : null,
  };
}

/**
 * The software list as objects, for a fresh record or one read back from
 * SharePoint, which holds only the raw lines.
 */
export function softwareOf(device) {
  if (Array.isArray(device?.installedSoftware)) return device.installedSoftware;
  return parseSoftware(linesOf(device?.installedSoftwareRaw));
}
