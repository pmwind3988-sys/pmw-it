import { isStale, labelOf, ATTENTION_LEVELS } from '../deviceFilters.js';

/**
 * A scan that failed is not a healthy machine — it is an unknown one. Counting
 * it would pull the fleet's average RAM down and park a machine with no CPU at
 * the top of the "all clear" list, so every figure but rescanNeeded ignores it.
 */
const complete = (devices) => devices.filter((device) => device.scanComplete !== false);

export function fleetSummary(devices, now = Date.now()) {
  const rows = complete(devices);
  const ram = rows.map((device) => device.installedRamGB).filter((n) => typeof n === 'number');

  return {
    total: rows.length,
    needsAttention: rows.filter(
      (device) => ATTENTION_LEVELS.includes(device.riskLevel),
    ).length,
    unsupportedOs: rows.filter((device) => device.osSupported === false).length,
    unprotected: rows.filter((device) => device.avProtected === false).length,
    avgRamGB: ram.length ? Math.round(ram.reduce((a, b) => a + b, 0) / ram.length) : null,
    staleScans: rows.filter((device) => isStale(device, now)).length,
  };
}

export function countBy(devices, keyFn) {
  const counts = new Map();

  for (const device of complete(devices)) {
    const label = labelOf(keyFn(device));
    counts.set(label, (counts.get(label) ?? 0) + 1);
  }

  return [...counts]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

export function scansByMonth(devices) {
  const counts = new Map();

  for (const device of complete(devices)) {
    if (typeof device.scannedOn !== 'number') continue;
    const date = new Date(device.scannedOn);
    const key = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }

  return [...counts]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, count]) => {
      const [year, month] = key.split('-');
      return { label: `${month}/${year}`, count };
    });
}

const AGE_RANK = { Obsolete: 0, Aging: 1, Unknown: 2, Current: 3 };

export function leaderboards(devices, now = Date.now()) {
  const rows = complete(devices);
  const withRam = rows.filter((device) => typeof device.installedRamGB === 'number');

  return {
    highestRam: [...withRam].sort((a, b) => b.installedRamGB - a.installedRamGB).slice(0, 5),
    lowestRam: [...withRam].sort((a, b) => a.installedRamGB - b.installedRamGB).slice(0, 5),
    oldest: [...rows]
      .sort((a, b) => (AGE_RANK[a.cpuAgeBand] ?? 2) - (AGE_RANK[b.cpuAgeBand] ?? 2))
      .slice(0, 5),
    recent: [...rows].sort((a, b) => (b.scannedOn ?? 0) - (a.scannedOn ?? 0)).slice(0, 5),
    // The cheap fix: a free slot means one stick, not a new machine.
    upgradeCandidates: rows.filter(
      (device) => device.ramUpgradable
        && typeof device.installedRamGB === 'number'
        && device.installedRamGB <= 8,
    ),
    rescanNeeded: devices.filter(
      (device) => device.scanComplete === false || isStale(device, now),
    ),
  };
}
