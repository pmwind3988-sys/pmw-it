import { describe, it, expect } from 'vitest';
import { applyOverrides, EMPTY_OVERRIDES } from './overrides.js';
import { UNSORTED_ID } from './buckets.js';

const raw = {
  columnName: 'Describe',
  settings: { threshold: 0.3, granularity: 0.45 },
  buckets: [
    { id: 'sap', label: 'SAP / ERP', description: 'd', hints: [] },
    { id: 'approvals', label: 'Approvals & Workflow', description: 'd', hints: [] },
  ],
  fragments: [
    { id: '0:0', row: 0, index: 0, text: 'SAP posting fails', severity: 0.6, bucketId: 'sap', score: 0.5, themeId: 't1' },
    { id: '1:0', row: 1, index: 0, text: 'Approvals are chased', severity: 0.4, bucketId: 'approvals', score: 0.6, themeId: 't1' },
    { id: '1:1', row: 1, index: 1, text: 'Nobody knows the status', severity: 0.2, bucketId: UNSORTED_ID, score: 0.1, themeId: 't2' },
  ],
  themes: [
    { id: 't1', name: 'sap · approval', fragmentIds: ['0:0', '1:0'] },
    { id: 't2', name: 'status', fragmentIds: ['1:1'] },
  ],
  oneOffIds: [],
  noIssueRows: [2],
};

describe('applyOverrides', () => {
  it('counts people, not fragments', () => {
    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    const approvals = analysis.buckets.find((b) => b.id === 'approvals');
    expect(approvals.count).toBe(1);
    const theme = analysis.themes.find((t) => t.id === 't1');
    expect(theme.count).toBe(2);
    expect(theme.respondents).toBe(2);
  });

  it('always offers Unsorted, even when nothing is in it', () => {
    const clean = { ...raw, fragments: raw.fragments.map((f) => ({ ...f, bucketId: 'sap' })) };
    const analysis = applyOverrides(clean, EMPTY_OVERRIDES);
    expect(analysis.buckets.some((b) => b.id === UNSORTED_ID)).toBe(true);
  });

  it('honours a hand retag', () => {
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, retags: { '1:1': 'approvals' } });
    expect(analysis.fragments.find((f) => f.id === '1:1').bucketId).toBe('approvals');
    expect(analysis.buckets.find((b) => b.id === 'approvals').count).toBe(2);
  });

  it('excludes noise from every count but keeps the row visible', () => {
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, noise: ['0:0'] });
    expect(analysis.fragments.find((f) => f.id === '0:0').noise).toBe(true);
    expect(analysis.buckets.find((b) => b.id === 'sap').count).toBe(0);
  });

  it('renames and merges themes', () => {
    const analysis = applyOverrides(raw, {
      ...EMPTY_OVERRIDES,
      themeNames: { t1: 'Approval chasing' },
      themeMerges: { t2: 't1' },
    });
    const merged = analysis.themes.find((t) => t.id === 't1');
    expect(merged.name).toBe('Approval chasing');
    expect(merged.count).toBe(3);
    expect(analysis.themes.some((t) => t.id === 't2')).toBe(false);
  });

  it('drops an override pointing at a fragment that no longer exists', () => {
    // A re-import with different data must not corrupt the screen.
    const analysis = applyOverrides(raw, { ...EMPTY_OVERRIDES, retags: { '99:9': 'sap' }, noise: ['99:9'] });
    expect(analysis.fragments).toHaveLength(3);
    expect(analysis.buckets.find((b) => b.id === 'sap').count).toBe(1);
  });

  it('survives a re-score: the retag is applied to the new raw result', () => {
    const overrides = { ...EMPTY_OVERRIDES, retags: { '1:1': 'sap' } };
    const rescored = { ...raw, fragments: raw.fragments.map((f) => ({ ...f, bucketId: 'approvals' })) };
    expect(applyOverrides(rescored, overrides).fragments.find((f) => f.id === '1:1').bucketId)
      .toBe('sap');
  });

  it('produces a priority order', () => {
    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    expect(analysis.priority.length).toBeGreaterThan(0);
    expect(analysis.priority[0]).toHaveProperty('score');
  });
});
