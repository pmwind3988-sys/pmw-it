import { describe, it, expect } from 'vitest';
import { splitResponseColumns, responseTableColumns } from './responseFields.js';

const forms = {
  columns: [
    { name: 'Id' },
    { name: 'Start time' },
    { name: 'Completion time' },
    { name: 'Email' },
    { name: 'Name' },
    { name: 'Which department are you in?' },
    { name: 'Describe the biggest issue you face' },
    { name: 'How often does it happen?' },
  ],
};

describe('splitResponseColumns', () => {
  const { identity, questions } = splitResponseColumns(forms);

  it('leads with the email address, then the time, then the department', () => {
    expect(identity.map((c) => c.key)).toEqual(['email', 'submitted', 'department']);
  });

  it('labels each identity field in plain words rather than the form\'s', () => {
    expect(identity.map((c) => c.label)).toEqual(['Email', 'Submitted', 'Department']);
    expect(identity[1].name).toBe('Start time');
  });

  it('takes one submission time even when the form exports two', () => {
    expect(identity.filter((c) => c.key === 'submitted')).toHaveLength(1);
    expect(questions.map((c) => c.name)).toContain('Completion time');
  });

  it('leaves every question in sheet order', () => {
    expect(questions.map((c) => c.name)).toEqual([
      'Id',
      'Completion time',
      'Name',
      'Describe the biggest issue you face',
      'How often does it happen?',
    ]);
  });

  it('returns a shorter list rather than a blank column when a field is missing', () => {
    const anonymous = { columns: [{ name: 'What went wrong?' }] };
    const result = splitResponseColumns(anonymous);
    expect(result.identity).toEqual([]);
    expect(result.questions).toHaveLength(1);
  });

  it('survives being asked before anything is imported', () => {
    expect(splitResponseColumns(null)).toEqual({ identity: [], questions: [] });
  });
});

describe('responseTableColumns', () => {
  it('caps the questions and never the identity fields', () => {
    const shown = responseTableColumns(forms, 1);
    expect(shown.map((c) => c.name)).toEqual([
      'Email', 'Start time', 'Which department are you in?', 'Id',
    ]);
  });

  it('shows every question when asked for all of them', () => {
    expect(responseTableColumns(forms, null)).toHaveLength(forms.columns.length);
  });
});
