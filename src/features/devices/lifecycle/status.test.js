import { describe, it, expect } from 'vitest';
import {
  STATUSES, statusOf, inFleet, IN_USE, IN_REPAIR, SPARE, RETIRED,
} from './status.js';

describe('statusOf', () => {
  it('reads a blank status as In use, so no row needs migrating', () => {
    expect(statusOf({})).toBe(IN_USE);
    expect(statusOf({ status: '' })).toBe(IN_USE);
    expect(statusOf({ status: 'Nonsense' })).toBe(IN_USE);
  });

  it('keeps a real status', () => {
    expect(statusOf({ status: 'Retired' })).toBe(RETIRED);
  });

  it('lists exactly four statuses', () => {
    expect(STATUSES).toEqual(['In use', 'In repair', 'Spare', 'Retired']);
  });
});

describe('inFleet', () => {
  it('counts machines in use and in repair, not spare or retired ones', () => {
    expect(inFleet({ status: IN_USE })).toBe(true);
    expect(inFleet({ status: IN_REPAIR })).toBe(true);
    expect(inFleet({})).toBe(true);
    expect(inFleet({ status: SPARE })).toBe(false);
    expect(inFleet({ status: RETIRED })).toBe(false);
  });
});
