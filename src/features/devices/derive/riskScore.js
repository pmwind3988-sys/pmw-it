/**
 * Additive and explainable on purpose: the dashboard shows WHY a machine
 * scored what it did, so every charge records its reason.
 *
 * Signals the scan could not read charge nothing. An unknown antivirus state
 * is not evidence of an unprotected machine, and charging for it would push
 * every partially-readable report into the attention queue.
 */
const RULES = [
  {
    points: 40,
    reason: 'Windows 10 or older — no security updates since 14 Oct 2025',
    applies: (device) => device.osSupported === false,
  },
  {
    points: 30,
    // Two different failures, same weight, never charged twice. "Nothing
    // enabled" is the worse one but it is rare; "Norton missing while Defender
    // holds the fort" is the common one, and it is the whole reason the scan
    // report carries an `Antivirus status` line of its own.
    reason: (device) =>
      (device.antivirusStatus !== 'Unknown' && !device.avProtected)
        ? 'No antivirus enabled at all'
        : 'Managed antivirus not installed or deactivated',
    applies: (device) => {
      if (device.antivirusStatus === 'Unknown') return false;
      if (!device.avProtected) return true;
      return device.antivirusStatus === 'Not Installed'
        || device.antivirusStatus === 'Installed — Inactive';
    },
  },
  {
    points: 25,
    reason: '4 GB of RAM or less',
    applies: (device) =>
      typeof device.installedRamGB === 'number' && device.installedRamGB <= 4,
  },
  {
    points: 15,
    reason: '8 GB of RAM or less',
    applies: (device) =>
      typeof device.installedRamGB === 'number'
      && device.installedRamGB > 4
      && device.installedRamGB <= 8,
  },
  { points: 25, reason: 'Obsolete processor', applies: (device) => device.cpuAgeBand === 'Obsolete' },
  { points: 10, reason: 'Aging processor', applies: (device) => device.cpuAgeBand === 'Aging' },
  { points: 10, reason: 'Mechanical hard disk', applies: (device) => device.hasHdd === true },
];

/**
 * Watch starts at 15, not 20, because 15 is the smallest charge among the four
 * signals IT asked to be flagged (unsupported OS 40, no antivirus 30, 8 GB or
 * less 15, obsolete CPU 25). At 20 a plain 8 GB machine scores OK and the RAM
 * signal never appears in the risk mix at all.
 *
 * The two smaller charges — a mechanical disk and an aging CPU, both 10 — are
 * deliberately left below the line: on their own they are worth noticing in a
 * leaderboard, not worth queueing for attention.
 */
function levelFor(score) {
  if (score >= 60) return 'Critical';
  if (score >= 40) return 'High';
  if (score >= 15) return 'Watch';
  return 'OK';
}

export function riskScore(device) {
  // An unscanned machine is unknown, not healthy. Scoring it zero would let a
  // failed scan sit at the top of the "all clear" list.
  if (device.scanComplete === false) {
    return {
      riskScore: null,
      riskLevel: 'Unknown',
      riskReasons: ['Scan incomplete — re-run the report'],
    };
  }

  const hits = RULES.filter((rule) => rule.applies(device));
  const score = hits.reduce((total, rule) => total + rule.points, 0);

  return {
    riskScore: score,
    riskLevel: levelFor(score),
    riskReasons: hits.map((rule) =>
      (typeof rule.reason === 'function' ? rule.reason(device) : rule.reason)),
  };
}
