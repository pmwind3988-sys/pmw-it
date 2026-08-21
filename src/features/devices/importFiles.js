import { deriveDevice } from './derive/deriveDevice.js';
import { parseReport } from './parse/parseReport.js';

export function readTextFile(file) {
  return file.text();
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

export async function importFiles(files) {
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

    devices.push(deriveDevice({ text, fileName: file.name, lastModified: file.lastModified }));
  }

  return dedupe(devices, rejected);
}
