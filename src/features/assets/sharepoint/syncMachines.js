import { readAllAssets } from './readAssets.js';
import { readAllDevices } from '../../devices/sharepoint/readDevices.js';
import { performLifecycle } from '../../devices/sharepoint/writeLifecycle.js';
import { machineFor, assetFor, machineMove } from '../machineLink.js';

/**
 * After a handover, a return or an undo: bring every machine in the device
 * list that IS one of the touched register rows into line with the register
 * (`machineLink.js`). Both lists are re-read, because the register rows have
 * just changed and somebody may have moved the machine meanwhile --
 * `performLifecycle` refuses a machine whose status moved under it.
 *
 * Never throws. The handover is the record; a device list that would not
 * follow is reported, and the machine's page offers the same move again.
 */
export async function syncMachines({
  siteUrl, token, assetKeys, recordedBy = '', note = '', justReturned = false,
}) {
  const keys = new Set((assetKeys ?? []).filter(Boolean));
  if (!keys.size) return { moved: [], failed: [] };

  let register;
  let devices;
  try {
    [register, devices] = await Promise.all([readAllAssets(siteUrl, token), readAllDevices(siteUrl, token)]);
  } catch (thrown) {
    return { moved: [], failed: [{ device: null, reason: thrown.message || 'The device list could not be read' }] };
  }

  const moved = [];
  const failed = [];
  for (const asset of register.filter((row) => keys.has(row.assetKey))) {
    const link = machineFor(asset, devices);
    if (!link || assetFor(link.device, register) !== asset) continue;
    const move = machineMove(asset, link.device, { justReturned });
    if (!move) continue;
    try {
      await performLifecycle({
        siteUrl,
        token,
        deviceId: link.device.id,
        action: move.action,
        input: { ...move.input, note: note || null },
        expectedStatus: move.expectedStatus,
        recordedBy,
      });
      moved.push({ device: link.device, asset, move });
    } catch (thrown) {
      failed.push({ device: link.device, asset, reason: thrown.message || 'could not be saved' });
    }
  }
  return { moved, failed };
}
