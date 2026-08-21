import { describe, it, expect } from 'vitest';
import { fleetSummary, countBy, scansByMonth, leaderboards } from './deviceStats.js';

const NOW = Date.UTC(2026, 7, 21);
const DAY = 86_400_000;

const rows = [
  {
    computerName: 'CRIT', riskLevel: 'Critical', osSupported: false, avProtected: true,
    installedRamGB: 2, ramUpgradable: false, cpuAgeBand: 'Obsolete', deviceType: 'Desktop',
    scanComplete: true, scannedOn: NOW - 10 * DAY,
  },
  {
    computerName: 'OLD', riskLevel: 'High', osSupported: false, avProtected: false,
    installedRamGB: 8, ramUpgradable: true, cpuAgeBand: 'Aging', deviceType: 'Laptop',
    scanComplete: true, scannedOn: NOW - 200 * DAY,
  },
  {
    computerName: 'GOOD', riskLevel: 'OK', osSupported: true, avProtected: true,
    installedRamGB: 16, ramUpgradable: false, cpuAgeBand: 'Current', deviceType: 'Laptop',
    scanComplete: true, scannedOn: NOW - DAY,
  },
  {
    computerName: 'BROKEN', riskLevel: 'Unknown', osSupported: null, avProtected: false,
    installedRamGB: null, ramUpgradable: false, cpuAgeBand: 'Unknown', deviceType: 'Unknown',
    scanComplete: false, scannedOn: NOW - 2 * DAY,
  },
];

describe('fleetSummary', () => {
  const summary = fleetSummary(rows, NOW);

  it('counts only complete scans', () => {
    expect(summary.total).toBe(3);
  });

  it('counts Critical and High together as needing attention', () => {
    expect(summary.needsAttention).toBe(2);
  });

  it('counts unsupported operating systems', () => {
    expect(summary.unsupportedOs).toBe(2);
  });

  it('counts unprotected machines', () => {
    expect(summary.unprotected).toBe(1);
  });

  it('averages RAM over complete scans only', () => {
    // (2 + 8 + 16) / 3 — the failed scan must not pull it down
    expect(summary.avgRamGB).toBe(9);
  });

  it('counts scans older than 180 days as stale', () => {
    expect(summary.staleScans).toBe(1);
  });

  it('reports a null average rather than NaN for an empty fleet', () => {
    expect(fleetSummary([], NOW).avgRamGB).toBe(null);
  });
});

describe('countBy', () => {
  it('counts by a key, biggest group first, excluding incomplete scans', () => {
    expect(countBy(rows, (d) => d.deviceType)).toEqual([
      { label: 'Laptop', count: 2 },
      { label: 'Desktop', count: 1 },
    ]);
  });

  it('labels a missing value rather than dropping the row', () => {
    expect(countBy([{ scanComplete: true, department: null }], (d) => d.department))
      .toEqual([{ label: 'Unassigned', count: 1 }]);
  });
});

describe('scansByMonth', () => {
  it('groups by month in chronological order', () => {
    const result = scansByMonth([
      { scanComplete: true, scannedOn: Date.UTC(2026, 6, 3) },
      { scanComplete: true, scannedOn: Date.UTC(2026, 7, 1) },
      { scanComplete: true, scannedOn: Date.UTC(2026, 7, 20) },
    ]);
    expect(result).toEqual([
      { label: '07/2026', count: 1 },
      { label: '08/2026', count: 2 },
    ]);
  });

  it('skips a row with no scan date', () => {
    expect(scansByMonth([{ scanComplete: true, scannedOn: null }])).toEqual([]);
  });
});

describe('leaderboards', () => {
  const boards = leaderboards(rows, NOW);

  it('ranks the most and least RAM', () => {
    expect(boards.highestRam[0].computerName).toBe('GOOD');
    expect(boards.lowestRam[0].computerName).toBe('CRIT');
  });

  it('ranks the oldest hardware by CPU age band', () => {
    expect(boards.oldest[0].computerName).toBe('CRIT');
  });

  it('lists the newest scans first', () => {
    expect(boards.recent[0].computerName).toBe('GOOD');
  });

  it('lists only machines fixable with a stick as upgrade candidates', () => {
    expect(boards.upgradeCandidates.map((d) => d.computerName)).toEqual(['OLD']);
  });

  it('lists incomplete and stale scans as needing a re-scan', () => {
    expect(boards.rescanNeeded.map((d) => d.computerName).sort()).toEqual(['BROKEN', 'OLD']);
  });

  it('keeps the failed scan out of every other board', () => {
    for (const board of ['highestRam', 'lowestRam', 'oldest', 'recent', 'upgradeCandidates']) {
      expect(boards[board].some((d) => d.computerName === 'BROKEN')).toBe(false);
    }
  });
});
