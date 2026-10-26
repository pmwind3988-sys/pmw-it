import { hardwareIdsOf } from './derive/deriveHardwareIds.js';
import { dayLabel } from './dayLabel.js';

/**
 * The "Serial and part numbers" panel's rows, pure. Battery is not stored by
 * the scan, so it is not a row here.
 */

/** The six identity fields, in the order the panel shows them. A missing value is null. */
export function identityFieldsOf(device) {
  const bios = [device.biosVersion, dayLabel(device.biosDate)].filter(Boolean).join(' · ') || null;
  const chassis = device.chassis
    ? `${device.chassis}${device.assetTag ? ` · asset tag ${device.assetTag}` : ' · no asset tag'}`
    : null;
  return [
    { key: 'manufacturer', label: 'Manufacturer', value: device.manufacturer || null, mono: false },
    { key: 'productNumber', label: 'Product number', value: device.productNumber || null, mono: true },
    { key: 'serialNumber', label: 'Serial number', value: device.serialNumber || null, mono: true },
    { key: 'motherboardSerial', label: 'Motherboard serial', value: device.motherboardSerial || null, mono: true },
    { key: 'bios', label: 'BIOS', value: bios, mono: false },
    { key: 'chassis', label: 'Chassis', value: chassis, mono: false },
  ];
}

const size = (sizeGB) => (sizeGB ? `${sizeGB} GB` : '');

/**
 * One row per memory stick, drive and screen. A row a scan did not fill in
 * keeps its place with a dash, so a missing serial reads as missing rather
 * than as one fewer part.
 */
export function partRowsOf(device) {
  const { memoryModules, driveSerials, monitorDetails } = hardwareIdsOf(device);

  const memory = memoryModules.map((stick, index) => ({
    key: `memory-${index}`,
    part: 'Memory',
    what: [stick.maker, size(stick.sizeGB)].filter(Boolean).join(' ') || 'Memory module',
    partNumber: stick.partNumber || null,
    serial: stick.serial || null,
    where: stick.slot || null,
  }));

  const drives = driveSerials.map((drive, index) => ({
    key: `drive-${index}`,
    part: 'Drive',
    what: [drive.model, drive.bus].filter(Boolean).join(' · ') || 'Drive',
    partNumber: null,
    serial: drive.serial || null,
    where: drive.firmware ? `Firmware ${drive.firmware}` : null,
  }));

  const screens = monitorDetails.map((screen, index) => {
    const week = screen.madeWeek ? ` week ${screen.madeWeek}` : '';
    return {
      key: `screen-${index}`,
      part: 'Screen',
      what: [screen.maker, screen.name].filter(Boolean).join(' ') || 'Screen',
      partNumber: null,
      serial: screen.serial || null,
      where: screen.madeYear ? `Made ${screen.madeYear}${week}` : null,
    };
  });

  return [...memory, ...drives, ...screens];
}
