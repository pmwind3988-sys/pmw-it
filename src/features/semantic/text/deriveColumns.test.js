import { describe, it, expect } from 'vitest';
import {
  deriveColumns, DERIVED_HEADERS, NO_ISSUE_LABEL, DERIVED_OVERRIDES,
} from './deriveColumns.js';

const analysis = {
  buckets: [
    { id: 'sap', label: 'SAP / ERP' },
    { id: 'approvals', label: 'Approvals & Workflow' },
  ],
  themes: [{ id: 't1', name: 'approval · chase' }],
  fragments: [
    { id: '0:0', row: 0, text: 'a', severity: 0.6, bucketId: 'sap', themeId: 't1', noise: false },
    { id: '0:1', row: 0, text: 'b', severity: 0.2, bucketId: 'approvals', themeId: 't1', noise: false },
    { id: '1:0', row: 1, text: 'c', severity: 0.9, bucketId: 'approvals', themeId: 't1', noise: true },
  ],
};

describe('deriveColumns', () => {
  const { headers, columns } = deriveColumns(analysis, 3);
  const col = (name) => columns[headers.indexOf(name)];

  it('emits one value per row for every column', () => {
    expect(headers).toEqual(DERIVED_HEADERS);
    for (const column of columns) expect(column).toHaveLength(3);
  });

  it('gives a row the category of its worst issue', () => {
    expect(col('Issue category')[0]).toBe('SAP / ERP');
  });

  it('lists every category a row raised, semicolon-joined', () => {
    // Deliberately the multi-value shape, so the chart canvas can count
    // it by option instead of by combination.
    expect(col('Issue categories')[0]).toBe('SAP / ERP;Approvals & Workflow');
  });

  it('leaves a row with no issues clearly marked', () => {
    expect(col('Issue category')[2]).toBe(NO_ISSUE_LABEL);
    expect(col('Issue count')[2]).toBe(0);
    expect(col('Severity')[2]).toBe(0);
  });

  it('ignores fragments marked as noise', () => {
    // Row 1's only fragment is noise, so the row has no issues at all.
    expect(col('Issue category')[1]).toBe(NO_ISSUE_LABEL);
    expect(col('Issue count')[1]).toBe(0);
  });

  it('reports severity as a whole number out of a hundred', () => {
    expect(col('Severity')[0]).toBe(60);
  });
});

describe('DERIVED_OVERRIDES', () => {
  it('declares the categories column multi rather than leaving it to inference', () => {
    // Most respondents raise one category, so most cells carry no
    // separator and the multi-select heuristic declines -- correctly, on
    // the evidence it has. This module wrote the column, so it says so.
    expect(DERIVED_OVERRIDES['Issue categories']).toEqual({ type: 'multi' });
  });
});
