import { describe, it, expect } from 'vitest';
import { profileDataset } from '../profile/profileDataset.js';
import { detectTextColumns } from './detectTextColumns.js';

const grid = {
  headers: ['ID', 'Email', 'Department', 'Describe'],
  rows: [
    [1, 'a@x.com', 'IT', 'Financial data is collected from multiple Excel files before reporting.'],
    [2, 'b@x.com', 'Finance', 'The consolidation is manual, repetitive and prone to human error every month.'],
    [3, 'c@x.com', 'Logistics', 'Updates arrive over WhatsApp so important information is regularly missed.'],
    [4, 'd@x.com', 'Finance', 'Reports are rebuilt from scratch and version control is guesswork.'],
    [5, 'e@x.com', 'Sales', 'Approvals sit for days because nobody is told they are waiting.'],
    [6, 'f@x.com', 'QAQC', 'Numbers are retyped between two systems and the typos are only found later.'],
  ],
};

describe('detectTextColumns', () => {
  const found = detectTextColumns(profileDataset(grid), grid);

  it('picks the column people wrote in', () => {
    expect(found.map((c) => c.name)).toContain('Describe');
  });

  it('rejects identifiers, categories and short unique values', () => {
    expect(found.map((c) => c.name)).not.toContain('ID');
    expect(found.map((c) => c.name)).not.toContain('Department');
    // Emails are unique and text-typed but far too short to be prose.
    expect(found.map((c) => c.name)).not.toContain('Email');
  });

  it('returns nothing when a sheet has no prose in it', () => {
    const plain = { headers: ['Dept'], rows: [['IT'], ['Finance'], ['IT']] };
    expect(detectTextColumns(profileDataset(plain), plain)).toEqual([]);
  });
});

describe('a survey small enough that the profiler calls prose categorical', () => {
  // 42 responses is under the profiler's 50-distinct categorical rule,
  // so the written-answer column is typed `categorical`, not `text`.
  // Gating on the type would make the whole feature invisible on exactly
  // the file it was built for. This test pins that it does not.
  const small = {
    headers: ['Department', 'Describe'],
    rows: Array.from({ length: 42 }, (_, i) => ([
      ['IT', 'Finance', 'Logistics'][i % 3],
      `Response ${i}: the monthly consolidation is manual and takes several days to finish properly.`,
    ])),
  };

  it('still finds the written-answer column', () => {
    const profile = profileDataset(small);
    expect(profile.columns.find((c) => c.name === 'Describe').type).toBe('categorical');
    expect(detectTextColumns(profile, small).map((c) => c.name)).toEqual(['Describe']);
  });
});
