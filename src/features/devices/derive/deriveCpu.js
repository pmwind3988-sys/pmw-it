import { cleanValue } from '../parse/placeholders.js';

/**
 * Families with no generation to read and no future to argue about -- the
 * bottom of both vendors' ranges, plus everything AMD built before Zen.
 * `A8-7410`, `FX-8350` and an `Athlon II` are all pre-2017 parts whose model
 * numbers follow no scheme worth decoding; without this they fall through to
 * the RAM-type fallback and a DDR4 board alone would call them Aging.
 */
const OBSOLETE_FAMILIES =
  /pentium|celeron|atom|phenom|sempron|turion|athlon\s*(?:\(tm\)\s*)?(?:ii|x[24]|64)\b|\bfx[-(]|\ba\d{1,2}-\d{4}|\be[12]-\d/i;

/**
 * Not every AMD part says "AMD" first. A scan that reports the processor as
 * `Athlon(tm) II X2 240` or `FX(tm)-8350` is still an AMD machine, and reading
 * it as `Other` hides it from the vendor breakdown on the fleet dashboard.
 */
const AMD_FAMILIES =
  /\bamd\b|ryzen|athlon|radeon|epyc|threadripper|phenom|sempron|turion|\bfx[-(]/i;

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

/**
 * Every processor ends up on ONE scale — the Intel generation it is contemporary
 * with — so that "is this AMD machine older than that Intel one" has an answer.
 *
 * The AMD half of that scale is its Zen architecture, which is what a Ryzen
 * badge does NOT tell you: a Ryzen 5 7530U and a Ryzen 5 7640U are both "7000"
 * and are three years of architecture apart. `intel` below is the generation
 * each Zen shipped against — Zen 3 against 11th gen, Zen 4 against 13th — and
 * it is what makes the two vendors comparable at all.
 */
const ZEN = {
  zen1: { label: 'Zen', intel: 7 },
  zenPlus: { label: 'Zen+', intel: 8 },
  zen2: { label: 'Zen 2', intel: 10 },
  zen3: { label: 'Zen 3', intel: 11 },
  zen3Plus: { label: 'Zen 3+', intel: 12 },
  zen4: { label: 'Zen 4', intel: 13 },
  zen5: { label: 'Zen 5', intel: 15 },
};

/**
 * AMD's 2022-and-later MOBILE numbering spells the architecture out in the
 * THIRD digit: 7*3*30U is Zen 3, 7*8*40U — third digit 4 — is Zen 4. It is the
 * only part of the model number that means anything about the age of the chip.
 */
const ARCH_DIGIT = {
  0: ZEN.zen1, 1: ZEN.zenPlus, 2: ZEN.zen2, 3: ZEN.zen3, 4: ZEN.zen4, 5: ZEN.zen5,
};

/** Desktop and APU parts, by series. 5000G is Zen 3, 8000G is Zen 4. */
const DESKTOP_SERIES = {
  1: ZEN.zen1, 2: ZEN.zenPlus, 3: ZEN.zen2, 4: ZEN.zen2, 5: ZEN.zen3, 7: ZEN.zen4, 9: ZEN.zen5,
};

/**
 * Mobile parts run one series behind their desktop namesakes — a 3500U is Zen+
 * where a desktop 3600 is Zen 2 — which is exactly the trap this table exists
 * to close. APUs (the `G` parts) follow the same lag.
 */
const MOBILE_SERIES = {
  1: ZEN.zen1, 2: ZEN.zen1, 3: ZEN.zenPlus, 4: ZEN.zen2, 5: ZEN.zen3,
  6: ZEN.zen3Plus, 7: ZEN.zen4, 8: ZEN.zen4, 9: ZEN.zen5,
};

const MOBILE_SUFFIX = /^(HX|HS|H|U|E|C)\b/i;
const APU_SUFFIX = /^GE?\b/i;

function ryzenArchitecture(sku, suffix) {
  const series = Number(sku[0]);
  const mobile = MOBILE_SUFFIX.test(suffix);

  // Only mobile numbering carries the architecture digit. A desktop 7950X is
  // Zen 4, and reading its third digit would call it Zen 5.
  if (mobile && sku.length === 4 && series >= 7) {
    return ARCH_DIGIT[Number(sku[2])] ?? MOBILE_SERIES[series] ?? null;
  }

  const table = mobile || APU_SUFFIX.test(suffix) ? MOBILE_SERIES : DESKTOP_SERIES;
  return table[series] ?? null;
}

function readGeneration(model) {
  // The scan usually writes it outright — no inference needed.
  const explicit = /(\d{1,2})(?:st|nd|rd|th)\s+Gen/i.exec(model);
  if (explicit) return { kind: 'intel', value: Number(explicit[1]) };

  const ultra = /Core\(TM\)\s+Ultra\s+\d+\s+(\d)\d{2}/i.exec(model);
  if (ultra) return { kind: 'ultra', value: Number(ultra[1]) };

  const core = /i[3579][- ](\d{4,5})/i.exec(model);
  if (core) return { kind: 'intel', value: intelGenerationFromSku(core[1]) };

  // "Ryzen AI 9 HX 370" and "Ryzen AI Max+ 395" are Zen 5 whatever their
  // three-digit number says, and they do not follow the four-digit scheme.
  if (/Ryzen\s+AI\b/i.test(model)) {
    return { kind: 'amd', value: null, arch: ZEN.zen5, series: null };
  }

  // What sits between "Ryzen" and the model number varies: a tier digit on a
  // laptop part, `Threadripper` on a workstation one, `PRO` on business
  // versions of either. The model number is the only part always present.
  const ryzen = /Ryzen\s+(?:Threadripper\s+)?(?:PRO\s+)?(?:\d\s+)?(?:PRO\s+)?(\d{4})([A-Z+]*)/i
    .exec(model);
  if (ryzen) {
    const [, sku, suffix = ''] = ryzen;
    return {
      kind: 'amd',
      value: Number(sku[0]),
      series: Number(sku[0]) * 1000,
      arch: ryzenArchitecture(sku, suffix),
    };
  }

  // Athlon did not stop at the pre-Zen parts above: the 3000G, the 3050U and
  // the Gold/Silver laptop chips are cut-down Zen, and belong on the scale
  // with the rest of them rather than in the "cannot tell" bucket.
  const athlon = /Athlon\s+(?:(?:Gold|Silver|PRO)\s+)*(\d{3,4})/i.exec(model);
  if (athlon) return { kind: 'amd', value: null, series: null, arch: ZEN.zen1 };

  return { kind: 'none', value: null };
}

/**
 * The one number every processor is ranked on: the Intel generation it stands
 * level with. Intel parts are themselves; Core Ultra 1 and 2 continue the
 * count at 14 and 15; a Ryzen is placed by its Zen architecture.
 */
function generationRank(generation) {
  if (generation.kind === 'intel') return generation.value ?? null;
  if (generation.kind === 'ultra') return generation.value ? 13 + generation.value : null;
  if (generation.kind === 'amd') return generation.arch?.intel ?? null;
  return null;
}

function readAgeBand(model, rank, ramType) {
  if (ramType && /^DDR[123]$/i.test(ramType)) return 'Obsolete';

  if (rank) {
    if (rank >= 10) return 'Current';
    if (rank >= 7) return 'Aging';
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
    return {
      cpuModel: null,
      cpuVendor: null,
      cpuGeneration: null,
      cpuArchitecture: null,
      cpuGenerationRank: null,
      cpuAgeBand: 'Unknown',
    };
  }

  let cpuVendor = 'Other';
  if (/intel/i.test(cpuModel)) cpuVendor = 'Intel';
  else if (AMD_FAMILIES.test(cpuModel)) cpuVendor = 'AMD';

  const generation = readGeneration(cpuModel);
  const cpuGenerationRank = generationRank(generation);

  let cpuGeneration = null;
  if (generation.kind === 'intel' && generation.value) cpuGeneration = String(generation.value);
  else if (generation.kind === 'ultra') cpuGeneration = `Ultra ${generation.value}`;
  else if (generation.kind === 'amd') {
    const series = generation.series
      ? `Ryzen ${generation.series}`
      : (/Athlon/i.test(cpuModel) ? 'Athlon' : 'Ryzen AI');
    cpuGeneration = generation.arch ? `${series} (${generation.arch.label})` : series;
  }

  return {
    cpuModel,
    cpuVendor,
    cpuGeneration,
    cpuArchitecture: generation.kind === 'amd' ? (generation.arch?.label ?? null) : null,
    cpuGenerationRank,
    cpuAgeBand: readAgeBand(cpuModel, cpuGenerationRank, ramType),
  };
}
