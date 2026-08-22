import { describe, it, expect } from 'vitest';
import {
  csvField, datasetToCsv, dashboardToJson, parseDashboardJson,
  DASHBOARD_EXPORT_VERSION,
} from './exporters.js';

const dataset = {
  rowCount: 3,
  columns: [
    {
      name: 'Dept',
      type: 'categorical',
      values: Int32Array.from([0, 1, -1]),
      dictionary: ['HR, and more', 'IT'],
    },
    {
      name: 'Amount',
      type: 'numeric',
      values: Float64Array.from([10, NaN, 30]),
      dictionary: null,
    },
    {
      name: 'Created',
      type: 'date',
      dateOnly: true,
      values: Float64Array.from([Date.UTC(2024, 0, 15), Date.UTC(2024, 1, 20), NaN]),
      dictionary: null,
    },
    {
      name: 'Active',
      type: 'boolean',
      values: Uint8Array.from([1, 0, 2]),
      dictionary: null,
    },
  ],
  byName: new Map([['Dept', 0], ['Amount', 1], ['Created', 2], ['Active', 3]]),
};

describe('csvField', () => {
  it('leaves a plain value unquoted', () => {
    expect(csvField('HR')).toBe('HR');
  });

  // Skipping this is how one address column silently shifts every later
  // column of a row by one.
  it.each([
    ['a,b', '"a,b"'],
    ['say "hi"', '"say ""hi"""'],
    ['line\nbreak', '"line\nbreak"'],
  ])('quotes %j', (input, expected) => {
    expect(csvField(input)).toBe(expected);
  });

  it('renders null and undefined as empty, not as the words', () => {
    expect(csvField(null)).toBe('');
    expect(csvField(undefined)).toBe('');
  });
});

describe('datasetToCsv', () => {
  const lines = () => datasetToCsv(dataset).split('\r\n');

  it('writes the column names as a header row', () => {
    expect(lines()[0]).toBe('Dept,Amount,Created,Active');
  });

  it('resolves dictionary codes back to their labels, quoting as needed', () => {
    expect(lines()[1].startsWith('"HR, and more"')).toBe(true);
  });

  it('renders dates as DD/MM/YYYY for a date-only column', () => {
    expect(lines()[1]).toContain('15/01/2024');
  });

  // A missing number is a gap, not a zero. Writing 0 here is how an
  // export quietly changes the answer to every average in the file.
  it('leaves missing values empty rather than writing a zero', () => {
    const second = lines()[2].split(',');
    expect(second[1]).toBe('');
  });

  it('leaves a missing category and a missing boolean empty too', () => {
    const third = lines()[3].split(',');
    expect(third[0]).toBe('');
    expect(third[3]).toBe('');
  });

  it('renders booleans as words rather than as 1 and 0', () => {
    // Asserted on the end of the line, not by splitting on commas: the
    // first field of row 1 is a QUOTED value containing a comma, so a
    // naive split misaligns -- which is the whole reason csvField
    // quotes it.
    expect(lines()[1].endsWith(',Yes')).toBe(true);
    expect(lines()[2].endsWith(',No')).toBe(true);
  });

  it('uses CRLF line endings, which is what Excel expects', () => {
    expect(datasetToCsv(dataset)).toContain('\r\n');
  });

  it('exports only the masked rows when given a mask', () => {
    const csv = datasetToCsv(dataset, Uint8Array.from([1, 0, 0]));
    expect(csv.split('\r\n')).toHaveLength(2); // header + one row
  });
});

describe('dashboard JSON', () => {
  const dashboard = {
    name: 'Ops',
    tiles: [{ id: 't1', chart: 'bar', title: 'X' }],
    globalFilters: [{ column: 'Dept', kind: 'in', values: ['IT'] }],
    datasetName: 'Requests',
  };

  it('round-trips through export and import', () => {
    const result = parseDashboardJson(dashboardToJson(dashboard, dataset), dataset);
    expect(result.ok).toBe(true);
    expect(result.dashboard.tiles).toEqual(dashboard.tiles);
    expect(result.dashboard.globalFilters).toEqual(dashboard.globalFilters);
  });

  it('carries the column list so an import can explain itself', () => {
    const parsed = JSON.parse(dashboardToJson(dashboard, dataset));
    expect(parsed.columns).toEqual(['Dept', 'Amount', 'Created', 'Active']);
    expect(parsed.version).toBe(DASHBOARD_EXPORT_VERSION);
  });

  it('names the columns this dataset is missing, without refusing to load', () => {
    const other = { columns: [{ name: 'Dept' }] };
    const result = parseDashboardJson(dashboardToJson(dashboard, dataset), other);
    expect(result.ok).toBe(true);
    expect(result.missingColumns).toEqual(['Amount', 'Created', 'Active']);
  });

  // A wrong file is a normal thing for someone to pick. The answer is a
  // sentence, not an exception.
  it('reports unreadable JSON rather than throwing', () => {
    expect(parseDashboardJson('not json at all', dataset)).toMatchObject({ ok: false });
  });

  it('refuses a version it does not understand, and says which', () => {
    const result = parseDashboardJson(JSON.stringify({ version: 99, tiles: [] }), dataset);
    expect(result.ok).toBe(false);
    expect(result.reason).toContain('99');
  });

  it('refuses a file with no tiles array', () => {
    const result = parseDashboardJson(
      JSON.stringify({ version: DASHBOARD_EXPORT_VERSION }), dataset,
    );
    expect(result.ok).toBe(false);
  });
});
