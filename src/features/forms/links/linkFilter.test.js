import { describe, it, expect } from 'vitest';
import {
  filterLinks, countLinks, readFilter, fromTill, linkItems, linkSerials,
  dayGroupLabel, linkTime, groupByDay, linkSteps, avatarTone,
} from './linkFilter.js';
import { LINK_STATUS } from './linkSchema.js';
import { OUT } from '../checklistForm.js';

// Local-time fixtures, so the day labels mean the same wherever the tests run.
const NOW = new Date(2026, 9, 9, 12).getTime();
const at = (daysAgo, hour = 10) => new Date(2026, 9, 9 - daysAgo, hour).toISOString();
const ms = (iso) => Date.parse(iso);
const tomorrow = new Date(NOW + 86400000).toISOString();

const links = [
  {
    id: 1,
    status: LINK_STATUS.WAITING,
    employeeName: 'Amir Hakim',
    formMode: 'In',
    created: at(0, 9),
    expiresOn: tomorrow,
    preset: { items: [{ item: 'Laptop', quantity: 1 }, { item: 'Mouse', quantity: 2 }] },
    handovers: { kind: 'issue', ids: [7] },
  },
  {
    id: 2,
    status: LINK_STATUS.SIGNED,
    employeeName: 'Siti Aminah',
    formMode: OUT,
    created: at(3),
    signedOn: at(1, 14),
    expiresOn: at(1),
    submitted: { serialNumbers: 'Laptop · Dell 5440 · S/N 5CD2\nABC123' },
    handovers: { kind: 'return', ids: [9] },
  },
  {
    id: 3,
    status: LINK_STATUS.SIGNED,
    employeeName: 'Ravi Kumar',
    formMode: 'In',
    created: at(10),
    signedOn: at(10, 11),
    expiresOn: tomorrow,
    preset: { checkedItems: ['Laptop', 'Mouse'] },
    handovers: null,
  },
  {
    id: 4,
    status: LINK_STATUS.WAITING,
    employeeName: 'Lim Wei',
    formMode: 'In',
    created: at(5),
    expiresOn: at(1),
    handovers: null,
  },
  {
    id: 5,
    status: LINK_STATUS.CANCELLED,
    employeeName: 'Nur Aisyah',
    formMode: OUT,
    created: at(2),
    expiresOn: tomorrow,
    handovers: { kind: 'issue', ids: [] },
  },
];
const ids = (list) => list.map((link) => link.id);

describe('the shared-checklists filter', () => {
  it('reads the query string, defaulting and ignoring what it does not know', () => {
    expect(readFilter(new URLSearchParams('show=signed&from=till&kind=return&q=amir')))
      .toEqual({ show: 'signed', source: 'till', kind: 'return', q: 'amir' });
    expect(readFilter(new URLSearchParams('show=bogus&from=x&kind=y')))
      .toEqual({ show: 'all', source: 'any', kind: 'any', q: '' });
    expect(readFilter(new URLSearchParams(''))).toEqual({ show: 'all', source: 'any', kind: 'any', q: '' });
  });

  it('shows a signed link as signed even after its expiry date', () => {
    expect(ids(filterLinks(links, { show: 'signed' }, NOW))).toEqual([2, 3]);
  });

  it('keeps expired and cancelled links out of Waiting, and puts them under Closed', () => {
    expect(ids(filterLinks(links, { show: 'waiting' }, NOW))).toEqual([1]);
    expect(ids(filterLinks(links, { show: 'closed' }, NOW))).toEqual([4, 5]);
  });

  it('keeps only links the till made, and only links it did not make', () => {
    expect(ids(filterLinks(links, { source: 'till' }, NOW))).toEqual([1, 2]);
    expect(ids(filterLinks(links, { source: 'share' }, NOW))).toEqual([3, 4, 5]);
    expect(fromTill(links[4])).toBe(false);
  });

  it('still honours the older till boolean', () => {
    expect(ids(filterLinks(links, { till: true }, NOW))).toEqual([1, 2]);
  });

  it('splits hand outs from returns by the OUT form type', () => {
    expect(ids(filterLinks(links, { kind: 'return' }, NOW))).toEqual([2, 5]);
    expect(ids(filterLinks(links, { kind: 'handout' }, NOW))).toEqual([1, 3, 4]);
  });

  it('matches the text query on name, form type, items and serials', () => {
    expect(ids(filterLinks(links, { q: 'AMIR' }, NOW))).toEqual([1]);
    expect(ids(filterLinks(links, { q: 'mouse' }, NOW))).toEqual([1, 3]);
    expect(ids(filterLinks(links, { q: '5cd2' }, NOW))).toEqual([2]);
    expect(ids(filterLinks(links, { q: 'dell' }, NOW))).toEqual([2]);
    expect(ids(filterLinks(links, { q: '   ' }, NOW))).toEqual([1, 2, 3, 4, 5]);
  });

  it('combines the filters', () => {
    expect(ids(filterLinks(links, { show: 'waiting', source: 'till' }, NOW))).toEqual([1]);
    expect(ids(filterLinks(links, { show: 'signed', kind: 'return' }, NOW))).toEqual([2]);
  });

  it('counts each status under the other filters, not under its own', () => {
    expect(countLinks(links, {}, NOW)).toEqual({ all: 5, waiting: 1, signed: 2, closed: 2 });
    expect(countLinks(links, { source: 'till' }, NOW)).toEqual({ all: 2, waiting: 1, signed: 1, closed: 0 });
    expect(countLinks(links, { kind: 'return', q: 'dell' }, NOW)).toEqual({ all: 1, waiting: 0, signed: 1, closed: 0 });
  });
});

describe('what a link covers', () => {
  it('reads named items with their quantities, else the checked list, else nothing', () => {
    expect(linkItems(links[0])).toEqual([{ name: 'Laptop', qty: 1 }, { name: 'Mouse', qty: 2 }]);
    expect(linkItems(links[2])).toEqual([{ name: 'Laptop', qty: 1 }, { name: 'Mouse', qty: 1 }]);
    expect(linkItems(links[3])).toEqual([]);
  });

  it('prefers what the employee signed over what IT pre-filled', () => {
    const signed = { preset: { items: [{ item: 'Laptop' }] }, submitted: { items: [{ item: 'Monitor', quantity: 3 }] } };
    expect(linkItems(signed)).toEqual([{ name: 'Monitor', qty: 3 }]);
  });

  it('splits serial lines into what they are and the serial', () => {
    expect(linkSerials(links[1])).toEqual([
      { what: 'Laptop · Dell 5440', serial: '5CD2' },
      { what: '', serial: 'ABC123' },
    ]);
    expect(linkSerials(links[0])).toEqual([]);
  });
});

describe('the timeline', () => {
  it('labels a moment as today, yesterday, this week, this month, or its month', () => {
    expect(dayGroupLabel(new Date(2026, 9, 9, 8).getTime(), NOW)).toBe('Today');
    expect(dayGroupLabel(new Date(2026, 9, 8, 23).getTime(), NOW)).toBe('Yesterday');
    expect(dayGroupLabel(new Date(2026, 9, 3, 9).getTime(), NOW)).toBe('Earlier this week');
    expect(dayGroupLabel(new Date(2026, 9, 2, 9).getTime(), NOW)).toBe('Earlier this month');
    expect(dayGroupLabel(new Date(2026, 8, 20, 9).getTime(), NOW)).toBe('September 2026');
  });

  it('says Undated for a moment it cannot read', () => {
    expect(dayGroupLabel('not a date', NOW)).toBe('Undated');
    expect(dayGroupLabel(undefined, NOW)).toBe('Undated');
  });

  it('dates a signed link by when it was signed, and any other by when it was sent', () => {
    expect(linkTime(links[1])).toBe(ms(at(1, 14)));
    expect(linkTime(links[0])).toBe(ms(at(0, 9)));
    expect(linkTime({ status: LINK_STATUS.SIGNED, signedOn: 'nope', created: at(4) })).toBe(ms(at(4)));
  });

  it('groups links by day, newest first', () => {
    const groups = groupByDay(links, NOW);
    expect(groups.map((group) => group.label)).toEqual(['Today', 'Yesterday', 'Earlier this week', 'September 2026']);
    expect(groups.flatMap((group) => ids(group.links)).sort()).toEqual([1, 2, 3, 4, 5]);
  });

  it('draws a link as the steps it has been through', () => {
    const sent = ms(at(0, 9));
    expect(linkSteps(links[0], NOW)).toEqual([
      { label: 'Sent', when: sent, done: true },
      { label: 'Handed over', when: sent, done: true },
      { label: 'Signed', when: null, done: false },
    ]);
    expect(linkSteps(links[1], NOW)).toEqual([
      { label: 'Sent', when: ms(at(3)), done: true },
      { label: 'Handed over', when: ms(at(3)), done: true },
      { label: 'Signed', when: ms(at(1, 14)), done: true },
    ]);
    expect(linkSteps(links[3], NOW)).toEqual([
      { label: 'Sent', when: ms(at(5)), done: true },
      { label: 'Expired', when: ms(at(1)), done: false },
    ]);
    expect(linkSteps(links[4], NOW)).toEqual([
      { label: 'Sent', when: ms(at(2)), done: true },
      { label: 'Cancelled', when: null, done: false },
    ]);
  });

  it('gives every name one of six avatar tones, the same each time', () => {
    for (const name of ['Amir Hakim', 'Siti Aminah', '', undefined]) {
      const tone = avatarTone(name);
      expect(Number.isInteger(tone)).toBe(true);
      expect(tone).toBeGreaterThanOrEqual(0);
      expect(tone).toBeLessThanOrEqual(5);
      expect(avatarTone(name)).toBe(tone);
    }
  });
});

describe('grouping a list sorted by when links were sent', () => {
  it('puts a link signed today under Today first, even if it was sent last week', () => {
    const DAY = 86400000;
    const sentRecently = { id: 'a', status: LINK_STATUS.WAITING, created: new Date(NOW - 3 * DAY).toISOString() };
    const signedToday = {
      id: 'b', status: LINK_STATUS.SIGNED,
      created: new Date(NOW - 6 * DAY).toISOString(), signedOn: new Date(NOW - 3600000).toISOString(),
    };
    const groups = groupByDay([sentRecently, signedToday], NOW);
    expect(groups.map((group) => group.label)).toEqual(['Today', 'Earlier this week']);
  });
});
