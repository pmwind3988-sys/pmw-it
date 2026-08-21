import { cleanValue } from '../parse/placeholders.js';

const OBSOLETE_FAMILIES = /pentium|celeron|atom/i;

/**
 * Intel 4-digit SKUs are ambiguous: i7-3667U is 3rd generation (first digit)
 * and i7-1355U is 13th (first two). A 4-digit number starting 10-14 is a
 * 10th-generation-or-later part; 2-9 is that generation.
 */
function intelGenerationFromSku(sku) {
  if (sku.length >= 5) return Number(sku.slice(0, 2));

  if (sku.length === 4) {
    const leading = Number(sku.slice(0, 2));
    return leading >= 10 && leading <= 14 ? leading : Number(sku[0]);
  }

  return null;
}

function readGeneration(model) {
  // The scan usually writes it outright — no inference needed.
  const explicit = /(\d{1,2})(?:st|nd|rd|th)\s+Gen/i.exec(model);
  if (explicit) return { kind: 'intel', value: Number(explicit[1]) };

  const ultra = /Core\(TM\)\s+Ultra\s+\d+\s+(\d)\d{2}/i.exec(model);
  if (ultra) return { kind: 'ultra', value: Number(ultra[1]) };

  const core = /i[3579][- ](\d{4,5})/i.exec(model);
  if (core) return { kind: 'intel', value: intelGenerationFromSku(core[1]) };

  const ryzen = /Ryzen\s+\d+\s+(\d)\d{3}/i.exec(model);
  if (ryzen) return { kind: 'amd', value: Number(ryzen[1]) };

  return { kind: 'none', value: null };
}

function readAgeBand(model, generation, ramType) {
  if (ramType && /^DDR[123]$/i.test(ramType)) return 'Obsolete';

  if (generation.kind === 'ultra') return 'Current';

  if (generation.kind === 'intel' && generation.value) {
    if (generation.value >= 10) return 'Current';
    if (generation.value >= 7) return 'Aging';
    return 'Obsolete';
  }

  if (generation.kind === 'amd' && generation.value) {
    // AMD mobile numbering does not map onto Intel generations — a 7430U is a
    // Zen 3 part wearing a 7000 badge — so it is ranked on its own series.
    if (generation.value >= 5) return 'Current';
    if (generation.value >= 3) return 'Aging';
    return 'Obsolete';
  }

  if (OBSOLETE_FAMILIES.test(model)) return 'Obsolete';
  if (ramType && /^DDR5$/i.test(ramType)) return 'Current';
  if (ramType && /^DDR4$/i.test(ramType)) return 'Aging';
  return 'Unknown';
}

export function deriveCpu(processorLines, ramType) {
  const cpuModel = processorLines.length ? cleanValue(processorLines[0]) : null;

  if (!cpuModel) {
    return { cpuModel: null, cpuVendor: null, cpuGeneration: null, cpuAgeBand: 'Unknown' };
  }

  let cpuVendor = 'Other';
  if (/intel/i.test(cpuModel)) cpuVendor = 'Intel';
  else if (/amd|ryzen/i.test(cpuModel)) cpuVendor = 'AMD';

  const generation = readGeneration(cpuModel);

  let cpuGeneration = null;
  if (generation.kind === 'intel' && generation.value) cpuGeneration = String(generation.value);
  else if (generation.kind === 'ultra') cpuGeneration = `Ultra ${generation.value}`;
  else if (generation.kind === 'amd') cpuGeneration = `Ryzen ${generation.value}000`;

  return {
    cpuModel,
    cpuVendor,
    cpuGeneration,
    cpuAgeBand: readAgeBand(cpuModel, generation, ramType),
  };
}
