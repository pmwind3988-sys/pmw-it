import { cleanValue } from './placeholders.js';

const split = (line) => line.split('|').map((part) => part.trim());

/** `Total Slots: 2 | Used Slots: 2` — a summary, not a stick. */
const SLOT_SUMMARY = /^total\s+slots\s*:/i;

export function parseSize(text) {
  const match = /(\d+(?:\.\d+)?)\s*(TB|GB|MB)/i.exec(text ?? '');
  if (!match) return null;

  const value = Number(match[1]);
  const unit = match[2].toUpperCase();
  if (unit === 'TB') return Math.round(value * 1024);
  if (unit === 'MB') return Math.round(value / 1024);
  return Math.round(value);
}

const parseInteger = (text) => {
  const match = /(\d+)/.exec(text ?? '');
  return match ? Number(match[1]) : null;
};

export function parseRamSlots(lines) {
  const sticks = [];
  let totalSlots = null;
  let usedSlots = null;

  for (const line of lines) {
    if (SLOT_SUMMARY.test(line)) {
      const [totalPart, usedPart = ''] = split(line);
      totalSlots = parseInteger(totalPart);
      usedSlots = parseInteger(usedPart);
      continue;
    }

    const [size, type, speed, vendor, partNumber] = split(line);
    sticks.push({
      sizeGB: parseSize(size),
      type: cleanValue(type),
      speedMhz: parseInteger(speed),
      vendor: cleanValue(vendor),
      partNumber: cleanValue(partNumber),
    });
  }

  // The scan leaves `Used Slots:` blank on 5 of 17 machines. Counting the
  // sticks it did report is strictly better than reporting nothing.
  if (usedSlots === null) usedSlots = sticks.length;

  return { sticks, totalSlots, usedSlots };
}

export function parseDrives(lines) {
  return lines.map((line) => {
    const [model, type, size] = split(line);
    // "Unspecified" means Win32_DiskDrive could not read MediaType. On every
    // machine in the sample set that is a spinning disk.
    const isSsd = /ssd/i.test(type ?? '');
    return {
      model: cleanValue(model),
      type: isSsd ? 'SSD' : 'HDD (assumed)',
      sizeGB: parseSize(size),
      mechanical: !isSsd,
    };
  });
}

const stripPrefix = (text, prefix) =>
  cleanValue((text ?? '').replace(new RegExp(`^${prefix}\\s*:\\s*`, 'i'), ''));

export function parseNetwork(lines) {
  if (!lines.length) return null;

  const [connection, ssid, ip, assignment] = split(lines[0]);
  return {
    connection: cleanValue(connection),
    ssid: stripPrefix(ssid, 'SSID'),
    ip: stripPrefix(ip, 'IP'),
    assignment: cleanValue(assignment),
  };
}

export function parseAntivirus(lines) {
  const byProduct = new Map();

  for (const line of lines) {
    const [product, state] = split(line);
    const name = cleanValue(product);
    if (!name) continue;

    const enabled = /enabled/i.test(state ?? '');
    // AMIR-HP lists HP Wolf Pro Security 22 times with conflicting states.
    // A product is protecting the machine if any of its entries is enabled.
    byProduct.set(name, (byProduct.get(name) ?? false) || enabled);
  }

  return [...byProduct].map(([product, enabled]) => ({ product, enabled }));
}

export function parsePairs(lines) {
  return lines.map((line) => {
    const [left, right = null] = split(line);
    return { left: cleanValue(left), right: right === null ? null : cleanValue(right) };
  });
}

export function parseOffice(lines) {
  return lines
    .flatMap((line) => line.split(','))
    .map(cleanValue)
    .filter(Boolean);
}

const REJECT_GPU = /virtualmonitordriver/i;
const REJECT_MONITOR = /^default monitor$/i;

export function parseGpus(lines) {
  return lines.map(cleanValue).filter((value) => value && !REJECT_GPU.test(value));
}

export function parseMonitors(lines) {
  return lines.map(cleanValue).filter((value) => value && !REJECT_MONITOR.test(value));
}

export function parseMailFiles(lines) {
  return parsePairs(lines).map(({ left, right }) => ({
    file: left,
    path: right,
    kind: /\.pst$/i.test(left ?? '') ? 'archive' : 'mailbox',
  }));
}
