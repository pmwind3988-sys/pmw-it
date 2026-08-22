import { describe, it, expect, vi } from 'vitest';
import { analyze } from './analysis.js';
import { applyOverrides, EMPTY_OVERRIDES } from './overrides.js';
import { deriveColumns, DERIVED_OVERRIDES } from './deriveColumns.js';
import { profileDataset } from '../profile/profileDataset.js';

// Verbatim shapes from the real export: the bracketed labels, the
// trailing semicolons, and the "no issue from IT" row.
const RESPONSES = [
  'no issue from IT ',
  'Financial data is currently collected from multiple Excel files and different subsidiaries. '
    + 'The process involves extensive manual consolidation, which is repetitive, time-consuming and prone to human error. '
    + 'Automating extraction and report generation would reduce turnaround time.',
  'Selected Challenge]: Data Collection\n[Detailed Description]:\n'
    + 'I need to collect and monitor information from multiple WhatsApp groups and Excel files. '
    + 'Because there are many different groups, important information can sometimes be missed.',
  'Approvals are chased by email and nobody knows the current status of a request. '
    + 'Reminders have to be sent manually every week.',
  'The monthly report is rebuilt from scratch each time and version control is guesswork.',
  'SAP postings fail when master data is wrong, and correcting it is a manual job.',
];

const fakeEmbed = vi.fn(async (texts) => texts.map((text) => {
  const lower = text.toLowerCase();
  return Float32Array.from([
    /approv|sign-off|status|remind|chase/.test(lower) ? 1 : 0,
    /sap|erp|posting|master data/.test(lower) ? 1 : 0,
    /consolidat|report|excel|file|version/.test(lower) ? 1 : 0,
    /whatsapp|group|message|missed|communicat/.test(lower) ? 1 : 0,
  ]);
}));

describe('the text analysis pipeline, end to end', () => {
  it('turns written answers into chartable columns', async () => {
    const raw = await analyze({
      texts: RESPONSES,
      breadths: [0, 0.9, 0.4, 0.3, 0.2, 0.2],
      buckets: [
        { id: 'approvals', label: 'Approvals & Workflow', description: 'approval sign-off status reminder chase', hints: [] },
        { id: 'sap', label: 'SAP / ERP', description: 'sap erp posting master data', hints: [] },
        { id: 'consolidation', label: 'Data Consolidation & Reporting', description: 'consolidating excel files into a report with version control', hints: [] },
        { id: 'communication', label: 'Communication & Coordination', description: 'whatsapp groups messages missed communication', hints: [] },
      ],
      columnName: 'Describe',
      embedAll: fakeEmbed,
    });

    // The non-answer is excluded, and the rest produced more issues than
    // there were respondents -- which is the whole point of splitting.
    expect(raw.noIssueRows).toContain(0);
    expect(raw.fragments.length).toBeGreaterThan(RESPONSES.length - 1);
    // Nothing carries the pasted-in label through.
    for (const fragment of raw.fragments) {
      expect(fragment.text).not.toContain('Detailed Description');
    }

    const analysis = applyOverrides(raw, EMPTY_OVERRIDES);
    const { headers, columns } = deriveColumns(analysis, RESPONSES.length);

    // The derived columns go through the ordinary profiler, and
    // "Issue categories" has to come out as a multi column or the
    // by-option chart in spec section 9 does not exist.
    const grid = {
      headers,
      rows: RESPONSES.map((_, r) => columns.map((column) => column[r])),
    };
    const profile = profileDataset(grid, DERIVED_OVERRIDES);

    expect(profile.columns.find((c) => c.name === 'Issue categories').type).toBe('multi');

    const severity = profile.columns.find((c) => c.name === 'Severity');
    expect(severity.type).toBe('numeric');
    expect(severity.role).toBe('measure');
  });
});
