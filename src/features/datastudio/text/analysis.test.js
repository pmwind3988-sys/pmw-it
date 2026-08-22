import { describe, it, expect, vi } from 'vitest';
import { analyze, buildFragments, MIN_FRAGMENTS_FOR_THEMES } from './analysis.js';
import { UNSORTED_ID } from './buckets.js';

// A deterministic stand-in for the model: one dimension per keyword.
// Nothing in this file loads or needs the real thing.
const fakeEmbed = vi.fn(async (texts) => texts.map((text) => {
  const lower = text.toLowerCase();
  return Float32Array.from([
    /approv|sign-off|status|remind|chase/.test(lower) ? 1 : 0,
    /sap|erp|posting|master data/.test(lower) ? 1 : 0,
  ]);
}));

const buckets = [
  { id: 'approvals', label: 'Approvals', description: 'approval sign-off status reminder chase', hints: [] },
  { id: 'sap', label: 'SAP', description: 'sap erp posting master data', hints: [] },
];

describe('buildFragments', () => {
  it('gives every fragment a stable id of row and position', () => {
    const fragments = buildFragments(
      ['One problem here that is long enough. And a second one, also long enough.'],
      [0],
    );
    expect(fragments.map((f) => f.id)).toEqual(['0:0', '0:1']);
  });

  it('skips a row that raised nothing', () => {
    expect(buildFragments(['no issue from IT', ''], [0, 0])).toEqual([]);
  });
});

describe('analyze', () => {
  const texts = [
    'Approvals sit for days and nobody sends a reminder about the status.',
    'We chase sign-off by email and the status is never visible to anyone.',
    'The SAP posting fails and master data has to be corrected by hand.',
    'ERP master data is wrong so every SAP posting needs manual repair.',
    'Chasing approval status wastes a whole afternoon every single week.',
    'no issue from IT',
  ];
  const breadths = [0.2, 0.4, 0.6, 0.2, 0.8, 0];

  it('files fragments into the buckets they belong to', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.fragments.filter((f) => f.bucketId === 'approvals').length).toBeGreaterThan(0);
    expect(raw.fragments.filter((f) => f.bucketId === 'sap').length).toBeGreaterThan(0);
  });

  it('records which rows raised nothing at all', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.noIssueRows).toContain(5);
  });

  it('discovers themes and names them', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.themes.length).toBeGreaterThan(0);
    for (const theme of raw.themes) expect(theme.name).not.toBe('');
  });

  it('skips theme discovery when there is almost nothing to group', async () => {
    const raw = await analyze({
      texts: ['Approvals sit for days and nobody sends a reminder.'],
      breadths: [0],
      buckets,
      embedAll: fakeEmbed,
    });
    expect(raw.fragments.length).toBeLessThan(MIN_FRAGMENTS_FOR_THEMES);
    expect(raw.themes).toEqual([]);
  });

  it('sends everything to Unsorted when no bucket matches', async () => {
    const raw = await analyze({
      texts, breadths, buckets, settings: { threshold: 1.01 }, embedAll: fakeEmbed,
    });
    expect(raw.fragments.every((f) => f.bucketId === UNSORTED_ID)).toBe(true);
  });

  it('reports progress as it goes', async () => {
    const onProgress = vi.fn();
    await analyze({ texts, breadths, buckets, embedAll: fakeEmbed, onProgress });
    expect(onProgress).toHaveBeenCalled();
    const stages = onProgress.mock.calls.map(([p]) => p.stage);
    expect(new Set(stages).size).toBeGreaterThan(1);
  });

  it('embeds the bucket descriptions, not the bucket names', async () => {
    fakeEmbed.mockClear();
    await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    const embedded = fakeEmbed.mock.calls.flatMap(([list]) => list);
    expect(embedded).toContain('approval sign-off status reminder chase');
    expect(embedded).not.toContain('Approvals');
  });

  it('returns an empty result rather than throwing on an empty sheet', async () => {
    const raw = await analyze({ texts: [], breadths: [], buckets, embedAll: fakeEmbed });
    expect(raw.fragments).toEqual([]);
    expect(raw.themes).toEqual([]);
  });

  it('keeps the vectors so a settings change never re-embeds', async () => {
    const raw = await analyze({ texts, breadths, buckets, embedAll: fakeEmbed });
    expect(raw.vectors).toHaveLength(raw.fragments.length);
  });
});
