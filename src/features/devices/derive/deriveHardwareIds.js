import { cleanValue, cleanSerial } from '../parse/placeholders.js';
import { splitParts, cellValue, linesOf, meaningful } from '../parse/parts.js';

/**
 * What a scan version 2 report says about the machine's identity: the maker's
 * own numbers, the BIOS, the chassis, and the three lists that carry one line
 * per part (memory sticks, drives, monitors).
 *
 * Each list is returned twice. As parsed objects, for the page to show. And as
 * the raw lines, because that is what is stored in SharePoint: a stored record
 * is read back with the raw text only, and `hardwareIdsOf` parses it again.
 *
 * Everything here reads `fields` from parseReport, so an old report (no such
 * blocks) gives nulls and empty lists rather than throwing.
 */

const DATE = /^\d{4}-\d{2}-\d{2}$/;
const SIZE = /^(\d+(?:\.\d+)?)\s*GB$/i;
const PART_NUMBER = /^P\/N\s*/i;
const SERIAL = /^S\/N\s*/i;
const FIRMWARE = /^Firmware\s*/i;
const MADE = /^Made\s+(\d{4})\s+week\s+(\d+)/i;
const ASSET_TAG = /^Asset tag:\s*/i;

const joinOrNull = (lines) => (lines.length ? lines.join('\n') : null);

export function parseMemoryModules(lines) {
  return meaningful(lines).map((line) => {
    const [slot, ...rest] = splitParts(line);
    const sizeCell = rest.find((cell) => SIZE.test(cell));
    // The maker is the one cell left once the size and the two numbers are set aside.
    const makerCell = rest.find((cell) => cell && !SIZE.test(cell)
      && !PART_NUMBER.test(cell) && !SERIAL.test(cell));

    return {
      slot: cleanValue(slot),
      sizeGB: sizeCell ? Number(SIZE.exec(sizeCell)[1]) : null,
      maker: cleanValue(makerCell),
      partNumber: cellValue(rest, PART_NUMBER),
      serial: cellValue(rest, SERIAL, cleanSerial),
    };
  });
}

export function parseDriveSerials(lines) {
  return meaningful(lines).map((line) => {
    const [model, ...rest] = splitParts(line);
    // The bus (NVMe, SATA) is the one bare cell after the model.
    const busCell = rest.find((cell) => cell && !SERIAL.test(cell) && !FIRMWARE.test(cell));

    return {
      model: cleanValue(model),
      serial: cellValue(rest, SERIAL, cleanSerial),
      bus: cleanValue(busCell),
      firmware: cellValue(rest, FIRMWARE),
    };
  });
}

export function parseMonitorDetails(lines) {
  return meaningful(lines).map((line) => {
    const [maker, ...rest] = splitParts(line);
    const madeCell = rest.find((cell) => MADE.test(cell));
    const made = madeCell ? MADE.exec(madeCell) : null;
    const nameCell = rest.find((cell) => cell && !SERIAL.test(cell) && !MADE.test(cell));

    return {
      maker: cleanValue(maker),
      name: cleanValue(nameCell),
      serial: cellValue(rest, SERIAL, cleanSerial),
      madeYear: made ? Number(made[1]) : null,
      // The scan writes `week 0` when the panel does not say; that is not a week.
      madeWeek: made && Number(made[2]) > 0 ? Number(made[2]) : null,
    };
  });
}

function parseBios(lines) {
  const [line] = meaningful(lines);
  if (!line) return { biosVersion: null, biosDate: null };

  const [vendor, version, date] = splitParts(line);
  const name = [cleanValue(vendor), cleanValue(version)].filter(Boolean).join(' ');
  return {
    biosVersion: name || null,
    biosDate: DATE.test(date ?? '') ? date : null,
  };
}

function parseChassis(lines) {
  const [line] = meaningful(lines);
  if (!line) return { chassis: null, assetTag: null };

  const cells = splitParts(line);
  return {
    chassis: cleanValue(cells[0]),
    assetTag: cellValue(cells, ASSET_TAG),
  };
}

/** `26H2 | Build 26300.9711 | 64-bit | Installed 2025-11-01` reads as `26H2 (build 26300.9711, 64-bit)`. */
function parseWindowsBuild(lines) {
  const [line] = meaningful(lines);
  if (!line) return null;

  const cells = splitParts(line);
  const display = cleanValue(cells[0]);
  const build = cellValue(cells, /^Build\s*/i);
  const arch = cleanValue(cells.find((cell) => /bit$/i.test(cell)));
  const detail = [build && `build ${build}`, arch].filter(Boolean).join(', ');

  if (!display) return detail || null;
  return detail ? `${display} (${detail})` : display;
}

export function deriveHardwareIds(fields) {
  const memoryModules = parseMemoryModules(fields['Memory Modules'] ?? []);
  const driveSerials = parseDriveSerials(fields['Drive Serials'] ?? []);
  const monitorDetails = parseMonitorDetails(fields['Monitor Details'] ?? []);

  return {
    manufacturer: cleanValue(fields.Manufacturer?.[0]),
    productNumber: cleanValue(fields['Product Number']?.[0]),
    modelFamily: cleanValue(fields['Model Family']?.[0]),
    motherboardSerial: cleanSerial(fields['Motherboard Serial']?.[0]),
    deviceUuid: cleanValue(fields['Device UUID']?.[0]),
    ...parseBios(fields.BIOS ?? []),
    ...parseChassis(fields.Chassis ?? []),
    windowsBuild: parseWindowsBuild(fields['Windows Build'] ?? []),

    memoryModules,
    memoryModulesRaw: joinOrNull(meaningful(fields['Memory Modules'])),
    driveSerials,
    driveSerialsRaw: joinOrNull(meaningful(fields['Drive Serials'])),
    monitorDetails,
    monitorDetailsRaw: joinOrNull(meaningful(fields['Monitor Details'])),
  };
}

/**
 * The three part lists as arrays, whether the record is fresh (arrays already
 * there) or read back from SharePoint (raw text only). The page calls this, not
 * the fields, so it does not care which kind of record it was given.
 */
export function hardwareIdsOf(device) {
  return {
    memoryModules: Array.isArray(device?.memoryModules)
      ? device.memoryModules : parseMemoryModules(linesOf(device?.memoryModulesRaw)),
    driveSerials: Array.isArray(device?.driveSerials)
      ? device.driveSerials : parseDriveSerials(linesOf(device?.driveSerialsRaw)),
    monitorDetails: Array.isArray(device?.monitorDetails)
      ? device.monitorDetails : parseMonitorDetails(linesOf(device?.monitorDetailsRaw)),
  };
}
