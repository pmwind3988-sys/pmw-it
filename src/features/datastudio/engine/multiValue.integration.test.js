import { describe, it, expect } from 'vitest';
import { profileDataset } from '../profile/profileDataset.js';
import { proposeCleanPlan } from '../clean/proposeCleanPlan.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';
import { aggregate } from './aggregate.js';

// The shape of the real survey export: a multi-select question whose
// answers are semicolon-joined with a trailing separator.
const HEADERS = ['Department', 'Which challenges', 'Describe'];
const ROWS = [
  ['IT', 'no issue;', 'no issue from IT '],
  ['Finance', 'Data Collection;Data Consolidation;Report Generation;', 'Financial data is collected from many files.'],
  ['Logistics', 'Data Collection;Approval Tracking;', 'Updates arrive on WhatsApp and get missed.'],
  ['Finance', 'Data Collection;Report Generation;', 'Reports are rebuilt by hand each month.'],
  ['Sales', 'Approval Tracking;', 'Approvals sit with nobody chasing them.'],
  ['QAQC', 'Data Collection;Data Consolidation;', 'Numbers are retyped between two systems.'],
];

describe('the multi-select path, end to end', () => {
  it('ranks the options a survey offered', () => {
    const grid = { headers: HEADERS, rows: ROWS };
    const profile = profileDataset(grid);

    const challenges = profile.columns.find((c) => c.name === 'Which challenges');
    expect(challenges.type).toBe('multi');

    const plan = proposeCleanPlan(profile, grid);
    const dataset = applyCleanPlan(grid, plan, profile);

    const result = aggregate(dataset, null, {
      encoding: { x: { column: 'Which challenges' }, y: [{ column: null, agg: 'count' }] },
      sort: { by: 'y', dir: 'desc' },
      limit: 10,
    });

    expect(result.categories[0]).toBe('Data Collection');
    expect(result.series[0].data[0]).toBe(4);
    // Six respondents, but more than six option-picks: that is the whole
    // point of the type.
    expect(result.series[0].data.reduce((a, b) => a + b, 0)).toBeGreaterThan(ROWS.length);
  });
});
