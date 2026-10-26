import { describe, it, expect } from 'vitest';
import { applyFilters, toCsv, ramBucket, isStale, labelOf, IN_FLEET } from './deviceFilters.js';

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
    department: '', osSupported: true, storageType: 'SSD only', installedRamGB: 8,
    avProtected: true, cpuAgeBand: 'Current',
  },
  {
    computerName: 'D', owner: 'Dee', riskLevel: 'High', deviceType: null,
    department: 'SALES', osSupported: true, storageType: null, installedRamGB: 16,
    avProtected: null, cpuAgeBand: null,
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
    expect(applyFilters(rows, {})).toHaveLength(4);
  });

  it('ignores a filter whose value is empty', () => {
    expect(applyFilters(rows, { risk: '' })).toHaveLength(4);
  });

  it('finds both bands behind the "Need attention" figure', () => {
    expect(applyFilters(rows, { attention: '1' }).map((r) => r.computerName)).toEqual(['A', 'D']);
  });

  it('matches the Unassigned label the charts count under', () => {
    expect(applyFilters(rows, { type: 'Unassigned' }).map((r) => r.computerName)).toEqual(['D']);
    expect(applyFilters(rows, { storage: 'Unassigned' }).map((r) => r.computerName)).toEqual(['D']);
    expect(applyFilters(rows, { cpu: 'Unassigned' }).map((r) => r.computerName)).toEqual(['D']);
  });

  it('leaves an unreadable antivirus state out of both buckets', () => {
    expect(applyFilters(rows, { av: 'Unprotected' }).map((r) => r.computerName)).toEqual(['A']);
    expect(applyFilters(rows, { av: 'Protected' }).map((r) => r.computerName)).toEqual(['B', 'C']);
  });

  it('filters by risk level', () => {
    expect(applyFilters(rows, { risk: 'Critical' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters by device type', () => {
    expect(applyFilters(rows, { type: 'Laptop' }).map((r) => r.computerName)).toEqual(['B', 'C']);
  });

  it('reads a blank as Unassigned', () => {
    expect(labelOf('')).toBe('Unassigned');
    expect(labelOf(null)).toBe('Unassigned');
    expect(labelOf('SALES')).toBe('SALES');
  });

  it('filters by department, matching a blank one as Unassigned', () => {
    expect(applyFilters(rows, { department: 'SALES' })).toHaveLength(2);
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
    expect(applyFilters(rows, { nonsense: 'x' })).toHaveLength(4);
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

describe('the persona filters', () => {
  const rows = [
    {
      computerName: 'A', personaLabel: 'Engineering / Technical / Media',
      licenseStatus: 'Unlicensed', serverDependent: true, networkRisk: 'Severe',
      formFactorMatches: true,
    },
    {
      computerName: 'B', personaLabel: 'Executive / Field',
      licenseStatus: 'Authentic', serverDependent: true, networkRisk: 'Fine',
      formFactorMatches: false,
    },
    {
      computerName: 'C', personaLabel: 'Logistics / Operations / Desk',
      licenseStatus: 'Authentic', serverDependent: false, networkRisk: 'None',
      formFactorMatches: true,
    },
  ];

  const names = (params) => applyFilters(rows, params).map((row) => row.computerName);

  it('finds every machine judged against one workload profile', () => {
    expect(names({ persona: 'Executive / Field' })).toEqual(['B']);
  });

  it('separates working off the server from being slowed by it', () => {
    expect(names({ server: 'Dependent' })).toEqual(['A', 'B']);
    expect(names({ server: 'Bottleneck' })).toEqual(['A']);
  });

  it('finds the licences that need looking at', () => {
    expect(names({ license: 'Unlicensed' })).toEqual(['A']);
  });

  it('finds the machines whose form factor does not suit the role', () => {
    expect(names({ formfit: '1' })).toEqual(['B']);
  });
});

describe('part filters', () => {
  const rows = [
    { computerName: 'A', parts: { ram: { grade: 'Critical' } }, criticalParts: ['ram'] },
    { computerName: 'B', parts: { ram: { grade: 'Optimal' } }, criticalParts: [] },
  ];
  it('filters by one part’s grade', () => {
    expect(applyFilters(rows, { part: 'ram:Critical' }).map((r) => r.computerName)).toEqual(['A']);
  });
  it('filters to machines with any critical part', () => {
    expect(applyFilters(rows, { critical: '1' }).map((r) => r.computerName)).toEqual(['A']);
  });
  it('ignores an old fit link instead of failing', () => {
    expect(applyFilters(rows, { fit: 'Critical' })).toHaveLength(2);
  });
});

describe('status and location filters', () => {
  const rows = [
    { computerName: 'A', status: null, location: 'F1' },
    { computerName: 'B', status: 'Retired', location: 'f1' },
    { computerName: 'C', status: 'Spare', location: null },
  ];

  it('filters by status, reading a blank as In use', () => {
    expect(applyFilters(rows, { status: 'In use' }).map((r) => r.computerName)).toEqual(['A']);
    expect(applyFilters(rows, { status: 'Retired' }).map((r) => r.computerName)).toEqual(['B']);
  });

  it('filters by location code, and by Unassigned for none', () => {
    expect(applyFilters(rows, { location: 'F1' }).map((r) => r.computerName)).toEqual(['A', 'B']);
    expect(applyFilters(rows, { location: 'Unassigned' }).map((r) => r.computerName)).toEqual(['C']);
  });

  it("filters by status 'In fleet', keeping blank and In use/In repair, dropping Spare and Retired", () => {
    expect(applyFilters(rows, { status: IN_FLEET }).map((r) => r.computerName)).toEqual(['A']);
  });
});

describe('risk filter by band', () => {
  const rows = [{ riskLevel: 'Critical' }, { riskLevel: 'High' }, { riskLevel: 'Watch' }, { riskLevel: 'OK' }];
  it('reads High as both stored top levels, and keeps old stored-value links working', () => {
    expect(applyFilters(rows, { risk: 'High' })).toHaveLength(2);
    expect(applyFilters(rows, { risk: 'Medium' })).toEqual([{ riskLevel: 'Watch' }]);
    expect(applyFilters(rows, { risk: 'Critical' })).toEqual([{ riskLevel: 'Critical' }]);
  });
});

describe('work filter and wider search', () => {
  it('finds machines needing work and searches serial and AnyDesk', () => {
    const rows = [
      { computerName: 'A', parts: { cpu: { grade: 'Critical' } }, serialNumber: '7H2KQ93' },
      { computerName: 'B', parts: { cpu: { grade: 'Optimal' } }, anydeskId: '1 284 553' },
    ];
    expect(applyFilters(rows, { work: '1' }).map((d) => d.computerName)).toEqual(['A']);
    expect(applyFilters(rows, { q: 'h2kq' }).map((d) => d.computerName)).toEqual(['A']);
    expect(applyFilters(rows, { q: '284' }).map((d) => d.computerName)).toEqual(['B']);
  });
});

describe('licence, software and scan filters', () => {
  const fleet = [
    { id: 1, scanVersion: '2', windowsActivation: 'Not activated', deviceJoin: 'Not joined', officePersonal: true,
      officeLicenseType: 'Microsoft 365', windowsLicenseType: 'OEM',
      installedSoftware: [{ name: 'AutoCAD 2024', version: '1', publisher: 'Autodesk', installedOn: null }] },
    { id: 2, scanVersion: '2 basic', windowsActivation: 'Activated', deviceJoin: 'Azure AD', officePersonal: false,
      windowsLicenseType: 'Retail' },
    // Old scan: the fields were never read, so nothing can be said about them.
    { id: 3, scanVersion: '1', windowsActivation: null, deviceJoin: null, officePersonal: null },
    { id: 4, scanVersion: null },
  ];
  const ids = (params) => applyFilters(fleet, params).map((d) => d.id);

  it('an old scan matches neither Not activated nor Not joined', () => {
    expect(ids({ winact: 'Not activated' })).toEqual([1]);
    expect(ids({ joined: 'Not joined' })).toEqual([1]);
  });

  it('flags a personal Office copy or account', () => {
    expect(ids({ officewarn: 'yes' })).toEqual([1]);
  });

  it('matches a licence by its stored choice', () => {
    expect(ids({ winlicence: 'OEM' })).toEqual([1]);
    expect(ids({ officelic: 'Microsoft 365' })).toEqual([1]);
  });

  it('narrows by a program name, case-insensitively', () => {
    expect(ids({ software: 'autocad' })).toEqual([1]);
    expect(ids({ software: 'excel' })).toEqual([]);
  });

  it('sorts a scan into new, basic or old, with a missing version as old', () => {
    expect(ids({ scan: 'new' })).toEqual([1]);
    expect(ids({ scan: 'basic' })).toEqual([2]);
    expect(ids({ scan: 'old' })).toEqual([3, 4]);
  });
});
