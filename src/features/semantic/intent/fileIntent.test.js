import { describe, it, expect } from 'vitest';
import { readFileIntent, keywordsOf, relevanceOf } from './fileIntent.js';

describe('readFileIntent', () => {
  it('reads a survey title through the export scaffolding', () => {
    const intent = readFileIntent('Pain Points & Issues Across Departments (Responses).xlsx');
    expect(intent.title).toBe('Pain Points & Issues Across Departments');
    expect(intent.kind).toBe('issues');
    expect(intent.textFirst).toBe(true);
    expect(intent.keywords).toContain('department');
  });

  it('strips download and version noise from the title', () => {
    const intent = readFileIntent('Copy of Employee Feedback v2 2026-08-23 (1).xlsx');
    expect(intent.title).toBe('Employee Feedback');
    expect(intent.kind).toBe('feedback');
  });

  it('falls back to the sheet tab when the file name says nothing', () => {
    const intent = readFileIntent('Book1.xlsx', 'Helpdesk Tickets');
    expect(intent.title).toBe('Helpdesk Tickets');
    expect(intent.kind).toBe('tickets');
  });

  it('keeps the file name when it is real, even with a noisy tab', () => {
    const intent = readFileIntent('Asset Register.xlsx', 'Form Responses 1');
    expect(intent.title).toBe('Asset Register');
    expect(intent.kind).toBe('inventory');
  });

  it('leaves acronyms as typed', () => {
    expect(readFileIntent('SAP issues by HR.xlsx').title).toBe('SAP Issues By HR');
  });

  it('does not auto-analyse a sheet whose title is not about writing', () => {
    const intent = readFileIntent('Device Inventory 2026.xlsx');
    expect(intent.textFirst).toBe(false);
    expect(intent.kind).toBe('inventory');
  });

  it('returns a usable shape for a title it cannot classify', () => {
    const intent = readFileIntent('Q3 Numbers.xlsx');
    expect(intent.kind).toBe('generic');
    expect(intent.textFirst).toBe(false);
    expect(intent.title).toBe('Q3 Numbers');
  });

  it('survives a missing name', () => {
    const intent = readFileIntent(undefined);
    expect(intent.title).toBe('');
    expect(intent.keywords).toEqual([]);
  });
});

describe('keywordsOf', () => {
  it('drops stopwords, short words and plurals', () => {
    expect(keywordsOf('The issues across all Departments')).toEqual(['issue', 'department']);
  });
});

describe('relevanceOf', () => {
  const keywords = keywordsOf('Pain Points Across Departments');

  it('scores a column the title names', () => {
    expect(relevanceOf('Department', keywords)).toBeGreaterThan(0);
  });

  it('scores an unrelated column at zero rather than negative', () => {
    expect(relevanceOf('Timestamp', keywords)).toBe(0);
  });

  it('scores nothing when the title carried no keywords', () => {
    expect(relevanceOf('Department', [])).toBe(0);
  });
});
