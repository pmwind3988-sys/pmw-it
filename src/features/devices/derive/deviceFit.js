import { personaFor } from './persona.js';

/**
 * How well a machine suits the person using it.
 *
 * This is deliberately a second opinion, not a replacement for the risk score.
 * The risk score asks "is this machine dangerous or falling over?" and answers
 * the same way for everybody. This asks "is it the right machine for THIS
 * desk?", and a 16 GB laptop can be excellent in Sales and thin in Engineering.
 *
 * Four levels, and every one of them shows its working: each rule that fires
 * writes the sentence the dashboard prints, so nobody has to take the label on
 * trust.
 */

export const FIT_LEVELS = ['Critical', 'Needs Attention', 'Moderate', 'Optimal'];

/** Rules that make a machine unfit for work today. Any one of them is enough. */
const CRITICAL_RULES = [
  {
    text: 'Windows is past end of support — no security updates',
    applies: (d) => d.osSupported === false,
  },
  {
    text: 'Boots from a mechanical hard disk — everything waits on it',
    applies: (d) => d.storageType === 'HDD only',
  },
  {
    text: (d, p) => `${d.installedRamGB} GB of memory, below the ${p.minRamGB} GB floor for ${p.label}`,
    applies: (d, p) => typeof d.installedRamGB === 'number' && d.installedRamGB < p.minRamGB,
  },
  {
    text: 'Processor is obsolete — no upgrade path left on this board',
    applies: (d) => d.cpuAgeBand === 'Obsolete',
  },
  {
    text: (d) => (d.licenseStatus === 'Unlicensed'
      ? 'Office is installed outside the company licence'
      : 'No Office licence could be established on this machine'),
    applies: (d) => d.licenseStatus === 'Unlicensed' || d.licenseStatus === 'Undefined',
  },
  {
    text: 'Works off the server over Wi-Fi — the link is the bottleneck',
    applies: (d) => d.serverDependent === true && d.networkRisk === 'Severe',
  },
];

/** Rules that make a machine workable but under-provisioned for the desk. */
const ATTENTION_RULES = [
  {
    text: (d, p) => `${d.installedRamGB} GB of memory, under the ${p.goodRamGB} GB this work wants`,
    applies: (d, p) => typeof d.installedRamGB === 'number'
      && d.installedRamGB >= p.minRamGB
      && d.installedRamGB < p.goodRamGB,
  },
  {
    text: 'Drawing work on processor graphics — no dedicated card fitted',
    applies: (d, p) => p.needsDedicatedGpu && d.dedicatedGpu === false,
  },
  {
    text: 'Still has a spinning disk alongside the SSD',
    applies: (d) => d.hasHdd === true && d.storageType !== 'HDD only',
  },
  {
    text: 'Processor is a generation or more behind what this work wants',
    applies: (d, p) => d.cpuAgeBand === 'Aging'
      || (typeof d.cpuGenerationRank === 'number' && d.cpuGenerationRank < p.goodCpuRank),
  },
  {
    // Only ever fires when a scan reports free space; today's report does not,
    // so the rule sits ready rather than guessing at a figure it cannot see.
    text: (d) => `Disk is ${d.storageUsedPct}% full`,
    applies: (d) => typeof d.storageUsedPct === 'number' && d.storageUsedPct > 85,
  },
  {
    text: 'Reaches the server over Wi-Fi — a dock and a cable would be quicker',
    applies: (d) => d.serverDependent === true && d.networkRisk === 'Wireless',
  },
  {
    text: 'Works off the server, but the scan did not report the network link',
    applies: (d) => d.serverDependent === true && d.networkRisk === 'Unknown',
  },
];

const say = (rule, device, persona) =>
  (typeof rule.text === 'function' ? rule.text(device, persona) : rule.text);

const fire = (rules, device, persona) =>
  rules.filter((rule) => rule.applies(device, persona)).map((rule) => say(rule, device, persona));

/**
 * The portability tag. An informational label, never a fault: a desktop in
 * Sales is a note for the next refresh, not something anybody has to fix today.
 */
function portability(device, persona) {
  if (!persona.prefers) {
    return {
      suggestedFormFactor: null,
      formFactorNote: 'No department on the record, so no form factor is suggested',
      formFactorMatches: null,
    };
  }

  const matches = device.deviceType === 'Unknown' ? null : device.deviceType === persona.prefers;

  const note = persona.prefers === 'Laptop'
    ? 'This role works away from the desk — a laptop suits it better'
    : 'Deskbound work with headroom to buy — a desktop gives more for the money';

  return {
    suggestedFormFactor: persona.prefers,
    formFactorNote: matches ? `Already a ${persona.prefers.toLowerCase()} — a good match for this role` : note,
    formFactorMatches: matches,
  };
}

/** The one sentence the register prints in its "Action" column. */
function actionFor(status, criticals, attentions) {
  if (status === 'Unknown') return 'Re-run the scan';
  if (status === 'Critical') return `Replace or repair now — ${criticals[0]}`;
  if (status === 'Needs Attention') return `Plan an upgrade — ${attentions[0]}`;
  if (status === 'Optimal') return 'Nothing to do';
  return 'Monitor at the next review';
}

export function deviceFit(device) {
  const persona = personaFor(device.department);
  const tag = portability(device, persona);

  if (device.scanComplete === false) {
    return {
      personaKey: persona.key,
      personaLabel: persona.label,
      personaBlurb: persona.blurb,
      fitStatus: 'Unknown',
      fitReasons: ['Scan incomplete — nothing to judge this machine on'],
      actionRequired: 'Re-run the scan',
      ...tag,
    };
  }

  const criticals = fire(CRITICAL_RULES, device, persona);
  const attentions = fire(ATTENTION_RULES, device, persona);

  let fitStatus = 'Moderate';
  if (criticals.length) fitStatus = 'Critical';
  else if (attentions.length) fitStatus = 'Needs Attention';
  else if (
    typeof device.installedRamGB === 'number'
    && device.installedRamGB >= persona.comfortRamGB
    && device.storageType === 'SSD only'
    && device.osSupported === true
    && device.licenseStatus === 'Authentic'
    && (!persona.needsDedicatedGpu || device.dedicatedGpu === true)
    && tag.formFactorMatches !== false
  ) {
    fitStatus = 'Optimal';
  }

  const reasons = criticals.length ? criticals : attentions;

  return {
    personaKey: persona.key,
    personaLabel: persona.label,
    personaBlurb: persona.blurb,
    fitStatus,
    fitReasons: reasons.length
      ? reasons
      : [fitStatus === 'Optimal'
        ? 'Comfortably ahead of what this desk needs'
        : 'Meets the baseline for this desk'],
    actionRequired: actionFor(fitStatus, criticals, attentions),
    ...tag,
  };
}
