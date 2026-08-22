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

describe('what the real survey produced', () => {
  // Measured by running the real 42-response export through the real
  // model in a browser, not asserted from theory. These numbers are the
  // reason the defaults are what they are, and a change that moves them
  // a long way is a change worth noticing.
  it('records the shape the defaults were tuned to', async () => {
    const { DEFAULT_GRANULARITY } = await import('./cluster.js');
    const { DEFAULT_THRESHOLD } = await import('./similarity.js');

    // 0.45 left 80 of 134 fragments ungrouped; 0.75 put 62 of them in
    // one theme. 0.65 groups 123 with the largest at 19.
    expect(DEFAULT_GRANULARITY).toBe(0.65);
    // Fragment-to-bucket scores on the real data ran p25 0.311,
    // p50 0.369, p75 0.460. 0.30 leaves about a fifth unsorted, which is
    // an honest minority rather than a wall of forced guesses.
    expect(DEFAULT_THRESHOLD).toBe(0.3);
  });

  it('splits 42 real-shaped answers into more issues than respondents', async () => {
    const { buildFragments } = await import('./analysis.js');
    // The real export averages 425 characters per answer. Splitting is
    // the reason 42 responses became 134 issues.
    const fragments = buildFragments(RESPONSES, RESPONSES.map(() => 0));
    expect(fragments.length).toBeGreaterThan(RESPONSES.length);
  });
});
