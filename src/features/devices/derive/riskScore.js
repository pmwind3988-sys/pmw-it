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
    reason: 'No active antivirus',
    applies: (device) => device.antivirusStatus !== 'Unknown' && !device.avProtected,
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

function levelFor(score) {
  if (score >= 60) return 'Critical';
  if (score >= 40) return 'High';
  if (score >= 20) return 'Watch';
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
    riskReasons: hits.map((rule) => rule.reason),
  };
}
