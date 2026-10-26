import { deriveDevice } from './derive/deriveDevice.js';
import { parseReport } from './parse/parseReport.js';

/**
 * Some reports are saved as UTF-16 (the scan's own writer does this on some
 * machines), and `file.text()` reads those as UTF-8 and garbles every line. A
 * byte-order mark says which encoding it is; without one, UTF-8 is assumed.
 */
export function decodeReport(bytes) {
  const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);

  if (view[0] === 0xff && view[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(view.subarray(2));
  }
  if (view[0] === 0xfe && view[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(view.subarray(2));
  }
  // The UTF-8 decoder drops a UTF-8 byte-order mark by itself.
  return new TextDecoder('utf-8').decode(view);
}

export async function readTextFile(file) {
  return decodeReport(await file.arrayBuffer());
}

function dedupe(devices, rejected) {
  const byName = new Map();

  for (const device of devices) {
    const key = (device.computerName ?? device.sourceFileName).toLowerCase();
    const existing = byName.get(key);

    if (!existing) {
      byName.set(key, device);
      continue;
    }

    const [keep, drop] = device.scannedOn > existing.scannedOn
      ? [device, existing]
      : [existing, device];

    byName.set(key, keep);
    rejected.push({
      fileName: drop.sourceFileName,
      reason: `Duplicate of ${keep.computerName} — kept the newer scan from ${keep.sourceFileName}`,
    });
  }

  return { devices: [...byName.values()], rejected };
}

export async function importFiles(files, { knownLocations } = {}) {
  const devices = [];
  const rejected = [];

  for (const file of files) {
    if (!/\.txt$/i.test(file.name)) {
      rejected.push({ fileName: file.name, reason: 'Not a .txt file' });
      continue;
    }

    let text;
    try {
      text = await readTextFile(file);
    } catch (error) {
      rejected.push({ fileName: file.name, reason: `Could not read the file: ${error.message}` });
      continue;
    }

    // Checked before deriving so an unrelated .txt is named as such rather
    // than imported as a machine with every field empty.
    if (!parseReport(text).isReport) {
      rejected.push({
        fileName: file.name,
        reason: 'Not a device report — no known fields found',
      });
      continue;
    }

    devices.push(deriveDevice({ text, fileName: file.name, lastModified: file.lastModified, knownLocations }));
  }

  return dedupe(devices, rejected);
}

/**
 * Fold a fresh import into one already under review, so a second drop adds to
 * the grid instead of replacing it. Runs the same duplicate rule across the
 * combined set: the newer scan of a machine wins, whichever drop it came from.
 */
export function mergeImports(existing, incoming) {
  const rejected = [...existing.rejected, ...incoming.rejected];
  return dedupe([...existing.devices, ...incoming.devices], rejected);
}
