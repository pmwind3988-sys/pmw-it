import { fromQueueItem, QueueTicket } from '../../ticketQueue/queue/ticket';
import {
  ages, arrivals, byArea, byPriority, changeText, closedAt, daysToClose, gaps, grainFor, headline, holders,
  leaderboard, median, percentile, slotText, speeds, splitCategory, stories, trend, trendWindow
} from './stats';

// Thursday 29 October 2026, 10:00 local.
const NOW = new Date(2026, 9, 29, 10);
let nextId = 1;
const at = (y: number, m: number, d: number, h = 10): string => new Date(y, m - 1, d, h).toISOString();
const P = (id: number, name: string): Record<string, unknown> => ({ Id: id, Title: name, EMail: `${name}@pmw` });

function t(over: Record<string, unknown>): QueueTicket {
  return fromQueueItem({
    Id: nextId++, Title: 'x', Status: 'New', Priority: 'Normal', Category: '', Department: '',
    Created: at(2026, 10, 1), Modified: at(2026, 10, 1), Cc: [], Assignedto0: P(7, 'Papathy'), ...over
  });
}

describe('basics', () => {
  it('median and percentile', () => {
    expect(median([5, 1, 3])).toBe(3);
    expect(median([1, 2, 3, 4])).toBe(2.5);
    expect(median([])).toBeUndefined();
    expect(percentile([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 0.9)).toBe(9);
  });

  it('a ticket closes on its completed date, else its last update; a cancellation is not a fix', () => {
    const done = t({ Status: 'Completed', Created: at(2026, 10, 1), Completed_x0020_Date: at(2026, 10, 4) });
    expect(daysToClose(done)).toBeCloseTo(3);
    const noDate = t({ Status: 'Completed', Modified: at(2026, 10, 2) });
    expect(closedAt(noDate)?.getDate()).toBe(2);
    expect(daysToClose(t({ Status: 'Cancelled', Modified: at(2026, 10, 2) }))).toBeUndefined();
    expect(closedAt(t({ Status: 'On Hold' }))).toBeUndefined();
  });
});

describe('trend', () => {
  it('picks days, weeks or months by the span', () => {
    expect(grainFor(new Date(2026, 9, 1), new Date(2026, 10, 1))).toBe('day');
    expect(grainFor(new Date(2026, 6, 1), new Date(2026, 10, 1))).toBe('week');
    expect(grainFor(new Date(2025, 10, 1), new Date(2026, 10, 1))).toBe('month');
  });

  it('"any time" shows the past 12 months, and stops at today', () => {
    const w = trendWindow({}, NOW);
    expect(w.from).toEqual(new Date(2025, 10, 1));
    expect(w.to).toEqual(new Date(2026, 9, 30));
    expect(trendWindow({ from: new Date(2026, 9, 1), to: new Date(2026, 10, 1) }, NOW).to).toEqual(new Date(2026, 9, 30));
  });

  it('counts raised, closed and what was still open at each month end', () => {
    const rows = [
      t({ Created: at(2026, 9, 10), Status: 'Completed', Completed_x0020_Date: at(2026, 10, 2) }),
      t({ Created: at(2026, 9, 20) }),
      t({ Created: at(2026, 10, 5) })
    ];
    const { grain, buckets } = trend(rows, {}, NOW);
    expect(grain).toBe('month');
    const sep = buckets.find((b) => b.start.getMonth() === 8);
    const oct = buckets.find((b) => b.start.getMonth() === 9);
    expect(sep?.label).toBe('Sep 26');
    expect([sep?.raised, sep?.closed, sep?.openAtEnd]).toEqual([2, 0, 2]);
    expect([oct?.raised, oct?.closed, oct?.openAtEnd]).toEqual([1, 1, 2]);
    expect(buckets[0].label).toBe('Nov 25');
  });
});

describe('headline', () => {
  const rows = [
    t({ Created: at(2026, 10, 3), Assignedto0: null }),
    t({ Created: at(2026, 9, 3), Status: 'Completed', Completed_x0020_Date: at(2026, 10, 3) }),
    t({ Created: at(2025, 10, 9), Priority: 'Critical' }),
    t({ Created: at(2026, 9, 1), Status: 'Completed', Completed_x0020_Date: at(2026, 9, 2) })
  ];
  const h = headline(rows, rows, NOW);

  it('counts open, unassigned and the oldest open ticket from the present', () => {
    expect(h.openNow).toBe(2);
    expect(h.unassignedOpen).toBe(1);
    expect(h.oldestOpen?.days).toBe(385);
  });

  it('compares this calendar month with last', () => {
    expect([h.raisedThisMonth, h.raisedLastMonth]).toEqual([1, 2]);
    expect([h.closedThisMonth, h.closedLastMonth]).toEqual([1, 1]);
    expect(h.criticalInView).toBe(1);
  });

  it('words the change', () => {
    expect(changeText(45, 35)).toBe('↑ 29% on last month');
    expect(changeText(30, 40)).toBe('↓ 25% on last month');
    expect(changeText(5, 0)).toBe('');
    expect(changeText(4, 4)).toBe('Same as last month');
  });
});

describe('people', () => {
  it('splits each person\'s open tickets by status, unassigned included', () => {
    const rows = [
      t({ Status: 'New' }), t({ Status: 'In progress' }), t({ Status: 'On Hold' }),
      t({ Assignedto0: null }), t({ Status: 'Completed', Completed_x0020_Date: at(2026, 10, 2) })
    ];
    const hs = holders(rows);
    expect(hs.map((h) => `${h.name}:${h.new}/${h.inProgress}/${h.onHold}`)).toEqual(['Papathy:1/1/1', 'Unassigned:1/0/0']);
  });

  it('credits closes this month to Completed by, else the assignee, and names the fastest with 3+', () => {
    const fast = (d: number): QueueTicket =>
      t({ Created: at(2026, 10, d), Status: 'Completed', Completed_x0020_Date: at(2026, 10, d, 14), Assignedto0: P(8, 'Ashraf') });
    const rows = [
      fast(5), fast(6), fast(7),
      t({ Created: at(2026, 10, 1), Status: 'Completed', Completed_x0020_Date: at(2026, 10, 9) }),
      t({ Created: at(2026, 10, 1), Status: 'Completed', Completed_x0020_Date: at(2026, 10, 9), CompletedBy: P(9, 'Joe') }),
      t({ Created: at(2026, 9, 1), Status: 'Completed', Completed_x0020_Date: at(2026, 9, 9) })
    ];
    const { closers, fastest } = leaderboard(rows, NOW);
    expect(closers.map((c) => `${c.name}:${c.closed}`)).toEqual(['Ashraf:3', 'Joe:1', 'Papathy:1']);
    expect(fastest?.name).toBe('Ashraf');
  });
});

describe('what and how urgent', () => {
  it('folds retired category names into their area', () => {
    expect(splitCategory('SAP - Change')).toEqual({ area: 'SAP', kind: 'Change' });
    expect(splitCategory('Access/Security - Request')).toEqual({ area: 'Access', kind: 'Request' });
    expect(splitCategory('Application')).toEqual({ area: 'Application/Software', kind: '' });
  });

  it('counts broken vs request per area, and the uncategorised apart', () => {
    const rows = [
      t({ Category: 'SAP - Incident' }), t({ Category: 'SAP - Request' }), t({ Category: 'SAP - Change' }),
      t({ Category: 'Software - Incident' }), t({ Category: '' })
    ];
    const { areas, uncategorised } = byArea(rows);
    expect(areas[0]).toMatchObject({ area: 'SAP', incident: 1, request: 1, other: 1, total: 3 });
    expect(areas[1].label).toBe('Software & apps');
    expect(uncategorised).toBe(1);
  });

  it('orders priorities Critical first and keeps the odd ones', () => {
    const rows = [t({ Priority: 'Low' }), t({ Priority: 'Critical' }), t({ Priority: 'Undefined' })];
    expect(byPriority(rows).map((p) => p.priority)).toEqual(['Critical', 'Low', 'Not set']);
  });
});

describe('when and how long', () => {
  it('puts each ticket in its weekday and hour, and finds the busiest', () => {
    const rows = [t({ Created: at(2026, 10, 29, 15) }), t({ Created: at(2026, 10, 22, 15) }), t({ Created: at(2026, 10, 26, 7) })];
    const heat = arrivals(rows);
    expect(heat.grid[3][8]).toBe(2);
    expect(heat.grid[0][0]).toBe(1);
    expect(heat.busiest).toEqual({ day: 3, col: 8, count: 2 });
    expect(heat.showWeekend).toBe(false);
    expect(slotText(3, 8)).toBe('Thursday, 3–4 PM');
    expect(slotText(0, 4)).toBe('Monday, 11–12 PM');
  });

  it('sorts closing times into buckets', () => {
    const c = (days: number): QueueTicket =>
      t({ Status: 'Completed', Created: at(2026, 9, 1), Completed_x0020_Date: new Date(new Date(2026, 8, 1, 10).getTime() + days * 864e5).toISOString() });
    expect(speeds([c(0.2), c(2), c(5), c(10), c(40)]).map((s) => s.count)).toEqual([1, 1, 1, 1, 1]);
  });

  it('ages only open tickets', () => {
    const rows = [t({ Created: at(2026, 10, 27) }), t({ Created: at(2025, 10, 9) }), t({ Created: at(2025, 1, 1), Status: 'Completed' })];
    expect(ages(rows, NOW).map((a) => a.count)).toEqual([1, 0, 0, 0, 1]);
  });
});

describe('gaps and stories', () => {
  it('measures what the records are missing', () => {
    const rows = [t({ Department: 'IT', Category: 'SAP - Incident' }), t({ Status: 'Completed' })];
    expect(gaps(rows).map((g) => `${g.have}/${g.of}`)).toEqual(['1/2', '1/2', '0/1']);
  });

  it('writes the three sentences', () => {
    const rows = [t({ Created: at(2026, 10, 29, 15), Status: 'Completed', Completed_x0020_Date: at(2026, 10, 29, 17) })];
    const s = stories(headline(rows, rows, NOW), arrivals(rows), NOW);
    expect(s[0].text).toBe('1 ticket in, 1 closed. The pile held steady.');
    expect(s[0].kicker).toBe('October so far');
    expect(s[1].text).toContain('Thursday, 3–4 PM');
    expect(s[2].text).toBe('Half of these tickets were closed within 2 hours.');
  });
});
