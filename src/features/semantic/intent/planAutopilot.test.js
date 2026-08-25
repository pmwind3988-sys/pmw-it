import { describe, it, expect } from 'vitest';
import { profileDataset } from '../profile/profileDataset.js';
import { detectTextColumns } from '../text/detectTextColumns.js';
import { planAutopilot, pickAnalyseColumn } from './planAutopilot.js';

// Shaped like the real survey export: bookkeeping on the left, one
// category, one measure, and the written answer people were asked for.
const grid = {
  headers: [
    'Timestamp', 'Email Address', 'Time taken', 'Department', 'Hours lost per week',
    'Describe the biggest issue you face',
  ],
  rows: [
    ['2026-08-01 09:00', 'a@x.com', 240, 'Finance', 6, 'Financial data is collected from multiple Excel files before any report can be produced.'],
    ['2026-08-01 09:05', 'b@x.com', 190, 'Finance', 4, 'The monthly consolidation is manual, repetitive and prone to human error every single time.'],
    ['2026-08-01 09:11', 'c@x.com', 320, 'Logistics', 8, 'Updates arrive over WhatsApp groups so important delivery information is regularly missed.'],
    ['2026-08-01 09:20', 'd@x.com', 150, 'Sales', 3, 'Approvals sit untouched for days because nobody is told that they are waiting on them.'],
    ['2026-08-01 09:31', 'e@x.com', 410, 'QAQC', 9, 'Numbers are retyped between two systems and the typos are only discovered much later.'],
    ['2026-08-01 09:44', 'f@x.com', 275, 'Logistics', 5, 'Finding the current version of a document means asking three people and hoping for the best.'],
  ],
};

const profile = profileDataset(grid);
const textColumns = detectTextColumns(profile, grid);

describe('planAutopilot', () => {
  const plan = planAutopilot({
    fileName: 'Pain Points & Issues Across Departments (Responses).xlsx',
    sheetName: 'Form Responses 1',
    profile,
    textColumns,
  });

  it('reads the subject from the file name', () => {
    expect(plan.intent.title).toBe('Pain Points & Issues Across Departments');
    expect(plan.intent.kind).toBe('issues');
  });

  it('hides the bookkeeping columns and nothing else', () => {
    expect(plan.hidden.map((c) => c.name))
      .toEqual(['Timestamp', 'Email Address', 'Time taken']);
  });

  it('keeps the columns the survey asked about', () => {
    const names = plan.hidden.map((c) => c.name);
    expect(names).not.toContain('Department');
    expect(names).not.toContain('Hours lost per week');
  });

  it('picks the written answer to read', () => {
    expect(plan.analyseColumn).toBe('Describe the biggest issue you face');
  });

  it('passes the title keywords on as chart focus', () => {
    expect(plan.focus).toContain('department');
  });

  it('picks a written answer even when the title is not about writing', () => {
    const stock = planAutopilot({
      fileName: 'Device Inventory 2026.xlsx', profile, textColumns,
    });
    expect(stock.analyseColumn).toBe('Describe the biggest issue you face');
  });

  it('never hides a written-answer column, whatever its header', () => {
    const notes = {
      headers: ['Department', 'Email address'],
      rows: grid.rows.map((row) => [row[3], row[5]]),
    };
    const notesProfile = profileDataset(notes);
    const notesText = detectTextColumns(notesProfile, notes);
    const result = planAutopilot({
      fileName: 'Feedback.xlsx', profile: notesProfile, textColumns: notesText,
    });
    expect(result.hidden.map((c) => c.name)).not.toContain('Email address');
  });

  it('hides nothing rather than emptying the canvas', () => {
    // Every chartable column is bookkeeping, so hiding them all would
    // leave no chart to draw at all.
    const admin = {
      headers: ['Timestamp', 'Email Address'],
      rows: grid.rows.map((row) => [row[0], row[1]]),
    };
    const result = planAutopilot({
      fileName: 'Responses.xlsx', profile: profileDataset(admin), textColumns: [],
    });
    expect(result.hidden).toEqual([]);
  });

  it('survives an import with no profile at all', () => {
    const result = planAutopilot({ fileName: 'x.xlsx' });
    expect(result.hidden).toEqual([]);
    expect(result.analyseColumn).toBeNull();
  });
});

describe('pickAnalyseColumn', () => {
  const columns = [
    { name: 'Any other comments', index: 1 },
    { name: 'Biggest issue you face', index: 2 },
  ];

  it('keeps the longest-written column when the title matches neither', () => {
    expect(pickAnalyseColumn(columns, ['budget'])).toBe('Any other comments');
  });

  it('promotes the column the title names', () => {
    expect(pickAnalyseColumn(columns, ['issue'])).toBe('Biggest issue you face');
  });

  it('has nothing to pick from an empty list', () => {
    expect(pickAnalyseColumn([], ['issue'])).toBeNull();
  });
});
