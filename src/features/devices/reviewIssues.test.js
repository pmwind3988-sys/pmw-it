import { describe, it, expect } from 'vitest';
import { issuesFor, sortForReview } from './reviewIssues.js';

const clean = {
  computerName: 'PC1', scanComplete: true, deviceType: 'Laptop',
  ramDiscrepancy: false, unknownLabels: [], owner: 'Ali',
};

describe('issuesFor', () => {
  it('finds nothing wrong with a clean record', () => {
    expect(issuesFor(clean)).toEqual([]);
  });

  it('reports an incomplete scan', () => {
    expect(issuesFor({ ...clean, scanComplete: false }))
      .toEqual(['Scan incomplete — most fields are empty']);
  });

  it('reports an unresolved device type', () => {
    expect(issuesFor({ ...clean, deviceType: 'Unknown' }))
      .toEqual(['Device type could not be determined']);
  });

  it('explains the RAM discrepancy rather than just flagging it', () => {
    expect(issuesFor({ ...clean, ramDiscrepancy: true, installedRamGB: 16, reportedRamGB: 15 }))
      .toEqual(['Reports 15 GB usable of 16 GB installed — the GPU reserves the rest']);
  });

  it('reports an unknown label by name', () => {
    expect(issuesFor({ ...clean, unknownLabels: [{ label: 'BitLocker Status', value: 'On' }] }))
      .toEqual(['New field found in the report: BitLocker Status']);
  });

  it('reports a missing owner', () => {
    expect(issuesFor({ ...clean, owner: null })).toEqual(['No owner could be resolved']);
  });
});

describe('sortForReview', () => {
  it('puts rows with problems first, then sorts by name', () => {
    const rows = [
      { ...clean, computerName: 'BBB' },
      { ...clean, computerName: 'AAA' },
      { ...clean, computerName: 'ZZZ', scanComplete: false },
    ];
    expect(sortForReview(rows).map((r) => r.computerName)).toEqual(['ZZZ', 'AAA', 'BBB']);
  });

  it('does not mutate the input', () => {
    const rows = [{ ...clean, computerName: 'B' }, { ...clean, computerName: 'A' }];
    sortForReview(rows);
    expect(rows.map((r) => r.computerName)).toEqual(['B', 'A']);
  });
});
