import { describe, it, expect } from 'vitest';
import { severityOf } from './severity.js';

const context = { meanLength: 80, breadth: 0 };

describe('severityOf', () => {
  it('scores a strongly worded issue above a mild one', () => {
    const mild = severityOf('We update the sheet each week.', context);
    const strong = severityOf(
      'The manual reconciliation is time-consuming, repetitive and prone to error, and deadlines are constantly missed.',
      context,
    );
    expect(strong).toBeGreaterThan(mild);
  });

  it('rises with each intensity term, then saturates', () => {
    const one = severityOf('The process is manual.', context);
    const three = severityOf('The process is manual, repetitive and time-consuming.', context);
    const many = severityOf(
      'The manual, repetitive, time-consuming, tedious, error-prone rework causes constant delays and duplicate effort.',
      context,
    );
    expect(three).toBeGreaterThan(one);
    expect(many).toBeGreaterThanOrEqual(three);
  });

  it('counts how many challenges the respondent picked', () => {
    const text = 'Reports take a long time to prepare.';
    expect(severityOf(text, { meanLength: 80, breadth: 1 }))
      .toBeGreaterThan(severityOf(text, { meanLength: 80, breadth: 0 }));
  });

  it('stays inside 0 and 1 whatever it is given', () => {
    const extreme = severityOf(
      'MANUAL!!! repetitive time-consuming tedious duplicate rework chase constantly unable cannot difficult delay missed overlooked bottleneck!!!'.repeat(5),
      { meanLength: 10, breadth: 1 },
    );
    expect(extreme).toBeLessThanOrEqual(1);
    expect(severityOf('', context)).toBeGreaterThanOrEqual(0);
  });
});
