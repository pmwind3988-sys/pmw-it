import { fromQueueItem, QueueTicket } from '../../ticketQueue/queue/ticket';
import { Drill, drillKey, drillList, DrillSets } from './drill';
import {
  ages, arrivals, byArea, byPriority, gaps, headline, holders, leaderboard, periodOf, speeds, trend
} from './stats';

// Thursday 29 October 2026, 10:00 local.
const NOW = new Date(2026, 9, 29, 10);
let nextId = 1;
const at = (y: number, m: number, d: number, h = 10): string => new Date(y, m - 1, d, h).toISOString();
const P = (id: number, name: string): Record<string, unknown> => ({ Id: id, Title: name, EMail: `${name}@pmw` });

function t(over: Record<string, unknown>): QueueTicket {
  return fromQueueItem({
    Id: nextId++, Title: 'x', Status: 'New', Priority: 'Normal', Category: '', Department: 'QC',
    Created: at(2026, 10, 1), Modified: at(2026, 10, 1), Cc: [], Assignedto0: P(7, 'Papathy'), ...over
  });
}

const rows: QueueTicket[] = [
  t({ Created: at(2026, 10, 5, 9), Category: 'SAP - Incident', Priority: 'Critical' }),
  t({ Created: at(2026, 10, 6, 14), Category: 'SAP - Request', Status: 'In progress', Assignedto0: P(8, 'Iskandar') }),
  t({ Created: at(2026, 10, 7, 7), Category: 'Software - Incident', Status: 'On Hold', Assignedto0: null }),
  t({ Created: at(2026, 9, 2, 19), Category: 'Hardware - Request', Status: 'Completed', Completed_x0020_Date: at(2026, 10, 3), Root_x0020_Cause: 'Cable' }),
  t({ Created: at(2026, 10, 10), Category: 'Hardware - Incident', Status: 'Completed', Completed_x0020_Date: at(2026, 10, 10, 15), CompletedBy: P(8, 'Iskandar') }),
  t({ Created: at(2026, 10, 12), Category: 'Email', Status: 'Completed', Completed_x0020_Date: at(2026, 10, 20), Department: '' }),
  t({ Created: at(2026, 3, 1), Category: '', Priority: 'Weird', Status: 'New' }),
  t({ Created: at(2026, 6, 15), Category: 'Network - Incident', Status: 'Cancelled', Modified: at(2026, 7, 1) }),
  t({ Created: at(2026, 10, 28, 12), Category: 'Access/Security - Request', Priority: 'High', Status: 'New' })
];

const date = { preset: 'all' } as const;
const period = periodOf(date, NOW);
const { buckets } = trend(rows, {}, NOW);
const sets: DrillSets = { scope: rows, inView: rows, period, buckets, now: NOW, nameOf: (id) => `#${id}` };
const count = (d: Drill): number => drillList(d, sets).tickets.length;

describe('a tap lists exactly the tickets the chart counted', () => {
  it('headline figures', () => {
    const h = headline(rows, rows, period, NOW);
    expect(count({ kind: 'open' })).toBe(h.openNow);
    expect(count({ kind: 'raised' })).toBe(h.raised);
    expect(count({ kind: 'closed' })).toBe(h.closed);
    expect(count({ kind: 'critical' })).toBe(h.criticalInView);
  });

  it('every point on both trend charts', () => {
    buckets.forEach((b, index) => {
      expect(count({ kind: 'bucket', index, series: 'raised' })).toBe(b.raised);
      expect(count({ kind: 'bucket', index, series: 'closed' })).toBe(b.closed);
      expect(count({ kind: 'bucket', index, series: 'open' })).toBe(b.openAtEnd);
    });
    const tabs = drillList({ kind: 'bucket', index: buckets.length - 1, series: 'closed' }, sets).tabs;
    expect(tabs?.map((x) => x.label)).toEqual([
      `Raised (${buckets[buckets.length - 1].raised})`,
      `Closed (${buckets[buckets.length - 1].closed})`
    ]);
  });

  it('every person and every part of their bar', () => {
    holders(rows).forEach((h) => {
      expect(count({ kind: 'holder', id: h.id })).toBe(h.total);
      expect(count({ kind: 'holder', id: h.id, state: 'new' })).toBe(h.new);
      expect(count({ kind: 'holder', id: h.id, state: 'inProgress' })).toBe(h.inProgress);
      expect(count({ kind: 'holder', id: h.id, state: 'onHold' })).toBe(h.onHold);
    });
    expect(drillList({ kind: 'holder', id: 0 }, sets).title).toBe('Unassigned');
  });

  it('every leaderboard row', () => {
    const { closers } = leaderboard(rows, NOW);
    expect(closers.length).toBeGreaterThan(0);
    closers.forEach((c) => expect(count({ kind: 'closer', id: c.id })).toBe(c.closed));
  });

  it('every area segment, and the uncategorised', () => {
    const { areas, uncategorised } = byArea(rows);
    areas.forEach((a) => {
      expect(count({ kind: 'area', area: a.area, part: 'incident' })).toBe(a.incident);
      expect(count({ kind: 'area', area: a.area, part: 'request' })).toBe(a.request);
      expect(count({ kind: 'area', area: a.area, part: 'other' })).toBe(a.other);
      expect(count({ kind: 'area', area: a.area })).toBe(a.total);
    });
    expect(count({ kind: 'area', area: '' })).toBe(uncategorised);
    expect(drillList({ kind: 'area', area: 'SAP', part: 'incident' }, sets).title).toBe("SAP · Something's broken");
  });

  it('every priority, every heat square, every speed and age bar, every gap', () => {
    byPriority(rows).forEach((p) => expect(count({ kind: 'priority', priority: p.priority })).toBe(p.count));
    const heat = arrivals(rows);
    heat.grid.forEach((row, day) => row.forEach((n, col) => expect(count({ kind: 'heat', day, col })).toBe(n)));
    speeds(rows).forEach((s, index) => expect(count({ kind: 'speed', index })).toBe(s.count));
    ages(rows, NOW).forEach((a, index) => expect(count({ kind: 'age', index })).toBe(a.count));
    gaps(rows).forEach((g, index) => expect(count({ kind: 'gap', index })).toBe(g.of - g.have));
  });
});

describe('order and keys', () => {
  it('lists newest first, the age bars oldest first, the speed bars slowest first', () => {
    const open = drillList({ kind: 'open' }, sets).tickets.map((x) => x.created);
    expect(open.slice().sort().reverse()).toEqual(open);
    const old = drillList({ kind: 'holder', id: 7 }, sets).tickets.map((x) => x.created);
    expect(old.slice().sort()).toEqual(old);
    expect(drillList({ kind: 'speed', index: 0 }, sets).order).toBe('slowest first');
  });

  it('gives the two halves of a trend point one key, and the open chart its own', () => {
    expect(drillKey({ kind: 'bucket', index: 3, series: 'raised' })).toBe(drillKey({ kind: 'bucket', index: 3, series: 'closed' }));
    expect(drillKey({ kind: 'bucket', index: 3, series: 'open' })).not.toBe(drillKey({ kind: 'bucket', index: 3, series: 'raised' }));
    expect(drillKey({ kind: 'area', area: 'SAP' })).not.toBe(drillKey({ kind: 'area', area: 'SAP', part: 'request' }));
  });
});
