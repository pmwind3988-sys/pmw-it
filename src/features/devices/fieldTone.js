/**
 * Which values on a device's page are the bad news and which are the good.
 *
 * The device page already states a risk level and lists its reasons in words.
 * This is the same judgement applied to the individual values that produced it,
 * so somebody reading down the page sees the 8 GB, the spinning disk and the
 * disabled antivirus without having to know which numbers are low.
 *
 * Only fields with a settled right answer are toned. A RAM discrepancy, a
 * static IP or a free memory slot are worth reading and are nobody's fault, so
 * they stay the colour of the rest of the page — colour that appears everywhere
 * says nothing.
 */

import { isItToolDrive } from './derive/itMedia.js';

const RISK = 'risk';
const OK = 'ok';

const band = (bad, good) => (value) => {
  if (bad(value)) return RISK;
  if (good(value)) return OK;
  return null;
};

const yesIsBad = band((v) => v === true, (v) => v === false);
const yesIsGood = band((v) => v === false, (v) => v === true);

/** Risk scores below the Watch line are the all-clear; see `riskScore.js`. */
const scoreTone = (value) => {
  if (typeof value !== 'number') return null;
  return value >= 15 ? RISK : OK;
};

const cpuTone = (device) => {
  if (device.cpuAgeBand === 'Current') return OK;
  if (device.cpuAgeBand === 'Aging' || device.cpuAgeBand === 'Obsolete') return RISK;
  return null;
};

const antivirusTone = (value) => {
  if (value === 'Active') return OK;
  if (value === 'Unknown' || value == null) return null;
  return RISK;
};

/** 8 GB or less is what the risk score charges for; 16 GB is a machine nobody
 *  needs to think about. Between the two, no opinion. */
const ramTone = (value) => {
  if (typeof value !== 'number') return null;
  if (value <= 8) return RISK;
  if (value >= 16) return OK;
  return null;
};

const FIELD_TONES = {
  riskScore: (device) => scoreTone(device.riskScore),
  riskLevel: (device) => (device.riskLevel === 'OK'
    ? OK
    : (['Critical', 'High', 'Watch'].includes(device.riskLevel) ? RISK : null)),
  riskReasons: () => RISK,
  scanComplete: (device) => yesIsGood(device.scanComplete),

  osSupported: (device) => yesIsGood(device.osSupported),
  windowsVersion: (device) => yesIsGood(device.osSupported),
  windowsMajor: (device) => yesIsGood(device.osSupported),

  antivirusStatus: (device) => antivirusTone(device.antivirusStatus),
  antivirusStatusRaw: (device) => antivirusTone(device.antivirusStatus),
  avProtected: (device) => yesIsGood(device.avProtected),

  cpuModel: cpuTone,
  cpuGeneration: cpuTone,
  cpuArchitecture: cpuTone,
  cpuGenerationRank: cpuTone,
  cpuAgeBand: cpuTone,

  installedRamGB: (device) => ramTone(device.installedRamGB),

  hasHdd: (device) => yesIsBad(device.hasHdd),
  storageType: (device) => {
    if (device.storageType === 'SSD only') return OK;
    if (device.storageType === 'Mixed' || device.storageType === 'HDD only') return RISK;
    return null;
  },
};

/** `'risk'`, `'ok'`, or nothing to say about this field. */
export function toneForField(device, key) {
  if (!device) return null;
  return FIELD_TONES[key]?.(device) ?? null;
}

const ENTRY_TONES = {
  riskReasons: () => RISK,
  antivirusProducts: (text) => (/\benabled\b/i.test(text) ? OK : RISK),
  // A drive line is `model | type | size`. `SSD` is the only type the scan
  // states outright; everything else it reports as `Unspecified`, which this
  // project reads as a spinning disk.
  storageDrivesRaw: (text) => {
    const [model] = text.split('|').map((part) => part.trim());
    // IT's own extraction disk is not this machine's problem, and colouring it
    // red would contradict the storage figures, which leave it out.
    if (isItToolDrive(model)) return null;
    return /\bssd\b/i.test(text) ? OK : RISK;
  },
};

/** The tone of ONE entry in a field that holds several — one antivirus
 *  product, one drive — where the field as a whole has no single answer. */
export function toneForEntry(key, text) {
  return ENTRY_TONES[key]?.(String(text ?? '')) ?? null;
}

/** Fields whose entries are toned one by one rather than all together. */
export const hasEntryTones = (key) => key in ENTRY_TONES;
