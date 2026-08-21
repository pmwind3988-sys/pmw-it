import { describe, it, expect } from 'vitest';
import { applyFilters, toCsv, ramBucket, isStale } from './deviceFilters.js';

const rows = [
  {
    computerName: 'A', owner: 'Ali', riskLevel: 'Critical', deviceType: 'Desktop',
    department: 'SALES', osSupported: false, storageType: 'Mixed', installedRamGB: 2,
    avProtected: false, cpuAgeBand: 'Obsolete',
  },
  {
    computerName: 'B', owner: 'Bea', riskLevel: 'OK', deviceType: 'Laptop',
    department: 'FINANCE', osSupported: true, storageType: 'SSD only', installedRamGB: 16,
    avProtected: true, cpuAgeBand: 'Current',
  },
  {
    computerName: 'C', owner: null, riskLevel: 'Watch', deviceType: 'Laptop',
    department: null, osSupported: true, storageType: 'SSD only', installedRamGB: 8,
    avProtected: true, cpuAgeBand: 'Current',
  },
];

describe('ramBucket', () => {
  it('buckets by installed size', () => {
    expect(ramBucket(2)).toBe('2 GB');
    expect(ramBucket(16)).toBe('16 GB');
    expect(ramBucket(null)).toBe('Unknown');
  });
});

describe('isStale', () => {
  const now = Date.UTC(2026, 7, 21);
  it('is stale past 180 days', () => {
    expect(isStale({ scannedOn: now - 200 * 86_400_000 }, now)).toBe(true);
    expect(isStale({ scannedOn: now - 10 * 86_400_000 }, now)).toBe(false);
  });

  it('is not stale when there is no scan date to judge', () => {
    expect(isStale({ scannedOn: null }, now)).toBe(false);
  });
});

describe('applyFilters', () => {
  it('returns everything when no filter is set', () => {
    expect(applyFilters(rows, {})).toHaveLength(3);
  });

  it('ignores a filter whose value is empty', () => {
    expect(applyFilters(rows, { risk: '' })).toHaveLength(3);
  });

  it('filters by risk level', () => {
    expect(applyFilters(rows, { risk: 'Critical' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters by device type', () => {
    expect(applyFilters(rows, { type: 'Laptop' }).map((r) => r.computerName)).toEqual(['B', 'C']);
  });

  it('filters by department, matching a blank one as Unassigned', () => {
    expect(applyFilters(rows, { department: 'SALES' })).toHaveLength(1);
    expect(applyFilters(rows, { department: 'Unassigned' }).map((r) => r.computerName))
      .toEqual(['C']);
  });

  it('filters unsupported operating systems', () => {
    expect(applyFilters(rows, { os: 'Unsupported' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters unprotected machines', () => {
    expect(applyFilters(rows, { av: 'Unprotected' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters by storage type, RAM bucket and CPU age', () => {
    expect(applyFilters(rows, { storage: 'SSD only' })).toHaveLength(2);
    expect(applyFilters(rows, { ram: '8 GB' }).map((r) => r.computerName)).toEqual(['C']);
    expect(applyFilters(rows, { cpu: 'Obsolete' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('searches computer name and owner, case-insensitively', () => {
    expect(applyFilters(rows, { q: 'bea' }).map((r) => r.computerName)).toEqual(['B']);
    expect(applyFilters(rows, { q: 'ALI' }).map((r) => r.computerName)).toEqual(['A']);
    // Substring, across both fields: "a" is in Ali and in Bea.
    expect(applyFilters(rows, { q: 'a' }).map((r) => r.computerName)).toEqual(['A', 'B']);
  });

  it('combines filters', () => {
    expect(applyFilters(rows, { department: 'SALES', risk: 'Critical' }).map((r) => r.computerName))
      .toEqual(['A']);
    expect(applyFilters(rows, { department: 'SALES', risk: 'OK' })).toHaveLength(0);
  });

  it('ignores an unrecognised filter key rather than returning nothing', () => {
    expect(applyFilters(rows, { nonsense: 'x' })).toHaveLength(3);
  });
});

describe('toCsv', () => {
  it('writes a header row and quotes what needs quoting', () => {
    const csv = toCsv(
      [{ computerName: 'A, Ltd', owner: 'He said "hi"' }],
      [{ key: 'computerName', label: 'Computer' }, { key: 'owner', label: 'Owner' }],
    );
    expect(csv).toBe('Computer,Owner\r\n"A, Ltd","He said ""hi"""');
  });

  it('renders a null as an empty cell', () => {
    expect(toCsv([{ owner: null }], [{ key: 'owner', label: 'Owner' }])).toBe('Owner\r\n');
  });

  it('joins an array cell with semicolons', () => {
    expect(toCsv([{ gpuList: ['Intel', 'NVIDIA'] }], [{ key: 'gpuList', label: 'GPU' }]))
      .toBe('GPU\r\nIntel; NVIDIA');
  });
});
