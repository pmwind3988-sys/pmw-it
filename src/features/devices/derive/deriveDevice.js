import { parseReport } from '../parse/parseReport.js';
import { cleanValue } from '../parse/placeholders.js';
import {
  parsePairs, parseNetwork, parseOffice, parseGpus, parseMonitors, parseMailFiles,
} from '../parse/parseValues.js';
import { deriveRam } from './deriveRam.js';
import { deriveStorage } from './deriveStorage.js';
import { deriveCpu } from './deriveCpu.js';
import { deriveIdentity } from './deriveIdentity.js';
import { deriveHealth } from './deriveHealth.js';
import { riskScore } from './riskScore.js';
import { enrichFit } from './enrichFit.js';

const firstOrNull = (lines) => (lines?.length ? cleanValue(lines[0]) : null);
const joinLines = (lines) => (lines?.length ? lines.join('\n') : null);

export function deriveDevice({ text, fileName, lastModified }) {
  const { fields, unknownLabels } = parseReport(text);

  const identity = deriveIdentity(fields, fileName);
  const ram = deriveRam(fields['RAM Slot Info'] ?? [], fields['Total RAM'] ?? []);
  const storage = deriveStorage(fields['Storage Drives'] ?? []);
  const cpu = deriveCpu(fields.Processor ?? [], ram.ramType);
  const health = deriveHealth(fields);

  const [motherboard] = parsePairs(fields.Motherboard ?? []);
  const network = parseNetwork(fields['Network Information'] ?? []);
  const mail = parseMailFiles(fields['Email data files found Active or Inactive account'] ?? []);
  const serverFolders = parsePairs(fields['Server folder'] ?? []);

  const base = {
    ...identity,
    ...ram,
    ...storage,
    ...cpu,
    ...health,

    computerModel: firstOrNull(fields['Computer Model']),
    motherboardVendor: motherboard?.left ?? null,
    motherboardModel: motherboard?.right ?? null,
    anydeskId: firstOrNull(fields.Anydesk),
    remarks: joinLines(fields.Remarks),

    scannedOn: lastModified,
    importedOn: Date.now(),
    sourceFileName: fileName,

    networkType: network?.connection ?? null,
    ssid: network?.ssid ?? null,
    ipAddress: network?.ip ?? null,
    ipAssignment: network?.assignment ?? 'Unknown',

    gpuList: parseGpus(fields.GPU ?? []),
    monitorCount: parseMonitors(fields.Monitor ?? []).length,
    monitorsRaw: joinLines(fields.Monitor),

    microsoftOffice: parseOffice(fields['Microsoft Office'] ?? []),
    adobeProducts: parsePairs(fields.Adobe ?? [])
      .filter((entry) => entry.left)
      .map((entry) => (entry.right ? `${entry.left} ${entry.right}` : entry.left)),
    mappedDrives: serverFolders.filter((entry) => entry.left).length,
    serverFolders: joinLines(fields['Server folder']),
    serverCredentials: joinLines(fields['PMW Server and credentials']),

    mailboxCount: mail.filter((entry) => entry.kind === 'mailbox').length,
    archiveCount: mail.filter((entry) => entry.kind === 'archive').length,
    emailDataFiles: joinLines(fields['Email data files found Active or Inactive account']),

    ramSlotInfoRaw: joinLines(fields['RAM Slot Info']),
    storageDrivesRaw: joinLines(fields['Storage Drives']),

    unknownLabels,
    rawReport: text,
  };

  return enrichFit({ ...base, ...riskScore(base) });
}
