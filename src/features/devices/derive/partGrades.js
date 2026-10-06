import { GRADES, UNKNOWN, PARTS, profileKeyFor } from '../standards/defaultStandard.js';

/**
 * Every part of a machine graded on its own, against the standard for the
 * desk it sits on. There is deliberately NO overall grade: a fast processor
 * does not make up for a hard disk, and one word for the whole machine hid
 * which part to fix.
 */
const RANK = Object.fromEntries(GRADES.map((grade, index) => [grade, index]));
const worse = (a, b) => {
  if (!a) return b;
  if (!b) return a;
  return RANK[a] <= RANK[b] ? a : b;
};

export function cutoffGrade(value, cut) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  if (value < cut.criticalBelow) return 'Critical';
  if (value < cut.attentionBelow) return 'Needs attention';
  if (value < cut.optimalFrom) return 'Moderate';
  return 'Optimal';
}

const ordinal = (n) => {
  const tens = n % 100;
  if (tens >= 11 && tens <= 13) return `${n}th`;
  return `${n}${({ 1: 'st', 2: 'nd', 3: 'rd' })[n % 10] ?? 'th'}`;
};
const gen = (n) => `${ordinal(n)} gen`;
const gb = (n) => `${n} GB`;

function reasonFor(grade, cut, unit) {
  if (grade === 'Critical') return `Below the ${unit(cut.criticalBelow)} floor`;
  if (grade === 'Needs attention') return `Under the ${unit(cut.attentionBelow)} needed for Moderate`;
  if (grade === 'Moderate') return `Under the ${unit(cut.optimalFrom)} needed for Optimal`;
  return `At or above the ${unit(cut.optimalFrom)} for Optimal`;
}

const unknown = (value, reason) => ({ grade: UNKNOWN, value, reason });

function gradeCpu(device, profile) {
  if (device.cpuAgeBand === 'Obsolete') {
    return { grade: 'Critical', value: device.cpuModel ?? 'Obsolete family', reason: 'Pentium, Celeron or AMD from before Ryzen — always Critical' };
  }
  const rank = device.cpuGenerationRank;
  const grade = cutoffGrade(rank, profile.cpu);
  if (!grade) return unknown(device.cpuModel ?? '—', 'The scan did not report a processor generation');
  return { grade, value: gen(rank), reason: reasonFor(grade, profile.cpu, gen) };
}

function gradeRam(device, profile) {
  const grade = cutoffGrade(device.installedRamGB, profile.ram);
  if (!grade) return unknown('—', 'The scan did not report the memory fitted');
  return { grade, value: gb(device.installedRamGB), reason: reasonFor(grade, profile.ram, gb) };
}

function gradeStorage(device, profile) {
  const typeGrade = profile.storageType[device.storageType] ?? null;
  const sizeGrade = cutoffGrade(device.storageTotalGB, profile.storageSize);
  const grade = worse(typeGrade, sizeGrade);
  if (!grade) return unknown('—', 'The scan did not report the disks');
  const value = [
    typeof device.storageTotalGB === 'number' ? gb(device.storageTotalGB) : null,
    typeGrade ? device.storageType : null,
  ].filter(Boolean).join(' · ');
  const typeDecides = typeGrade === grade && (!sizeGrade || RANK[typeGrade] <= RANK[sizeGrade]);
  const reason = typeDecides
    ? `${device.storageType} counts as ${grade} for this work`
    : reasonFor(sizeGrade, profile.storageSize, gb);
  return { grade, value, reason };
}

function gradeGraphics(device, profile) {
  if (device.dedicatedGpu === true) {
    const grade = profile.graphics.dedicated;
    return { grade, value: 'Dedicated card', reason: `A dedicated card counts as ${grade} for this work` };
  }
  if (device.dedicatedGpu === false) {
    const grade = profile.graphics.builtIn;
    return { grade, value: 'Built-in graphics', reason: `Built-in graphics count as ${grade} for this work` };
  }
  return unknown('—', 'The scan did not report the graphics adapter');
}

const WINDOWS_LABEL = {
  win11: 'Windows 11',
  win10Supported: 'Windows 10, still supported',
  outOfSupport: 'Out of support',
  other: 'An older or unrecognised Windows',
};

function windowsKind(device) {
  const text = device.windowsVersion ?? '';
  if (device.osSupported === false) return 'outOfSupport';
  if (device.windowsMajor === 11 || /windows 11/i.test(text)) return 'win11';
  if (device.windowsMajor === 10 || /windows 10/i.test(text)) return 'win10Supported';
  if (text || typeof device.windowsMajor === 'number') return 'other';
  return null;
}

function gradeWindows(device, profile) {
  const kind = windowsKind(device);
  if (!kind) return unknown('—', 'The scan did not report the Windows version');
  const grade = profile.windows[kind];
  return { grade, value: device.windowsVersion ?? WINDOWS_LABEL[kind], reason: `${WINDOWS_LABEL[kind]} counts as ${grade}` };
}

/** The portability tag, moved unchanged from the old deviceFit: a label, never a fault. */
function portability(device, profile) {
  if (!profile.prefers) {
    return { suggestedFormFactor: null, formFactorNote: 'No department on the record, so no form factor is suggested', formFactorMatches: null };
  }
  const matches = device.deviceType === 'Unknown' || !device.deviceType ? null : device.deviceType === profile.prefers;
  const note = profile.prefers === 'Laptop'
    ? 'This role works away from the desk — a laptop suits it better'
    : 'Deskbound work with headroom to buy — a desktop gives more for the money';
  return {
    suggestedFormFactor: profile.prefers,
    formFactorNote: matches ? `Already a ${profile.prefers.toLowerCase()} — a good match for this role` : note,
    formFactorMatches: matches,
  };
}

const labels = (keys) => keys.map((key) => PARTS.find((part) => part.key === key).label).join(', ');

export function partGrades(device, standard) {
  const personaKey = profileKeyFor(standard, device.department);
  const profile = standard.profiles[personaKey];
  const incomplete = device.scanComplete === false;

  const parts = incomplete
    ? Object.fromEntries(PARTS.map(({ key }) => [key, unknown('—', 'Scan incomplete — nothing to judge')]))
    : {
      cpu: gradeCpu(device, profile),
      ram: gradeRam(device, profile),
      storage: gradeStorage(device, profile),
      graphics: gradeGraphics(device, profile),
      windows: gradeWindows(device, profile),
    };

  const criticalParts = PARTS.filter(({ key }) => parts[key].grade === 'Critical').map(({ key }) => key);
  const attentionParts = PARTS.filter(({ key }) => parts[key].grade === 'Needs attention').map(({ key }) => key);

  let actionRequired = 'Nothing to do';
  if (incomplete) actionRequired = 'Re-run the scan';
  else if (criticalParts.length) actionRequired = `Upgrade now: ${labels(criticalParts)}`;
  else if (attentionParts.length) actionRequired = `Plan: ${labels(attentionParts)}`;

  return {
    personaKey,
    personaLabel: profile.label,
    personaBlurb: profile.blurb,
    ...portability(device, profile),
    parts,
    gradeCpu: parts.cpu.grade,
    gradeRam: parts.ram.grade,
    gradeStorage: parts.storage.grade,
    gradeGraphics: parts.graphics.grade,
    gradeWindows: parts.windows.grade,
    criticalParts,
    attentionParts,
    actionRequired,
  };
}
