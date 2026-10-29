import { rangeOf, inRange, describe as describeRange, parseDay } from './dateRange';
import { applyFilter, assigneesOf, defaultFilter, sortTickets, UNASSIGNED, QueueFilter } from './queueFilter';
import { fromQueueItem, htmlToText, QueueTicket } from './ticket';
import { chosenColumns, toCsv, safeText, exportFileName, ALL_COLUMN_KEYS, EXPORT_COLUMNS } from './exportColumns';

// Wednesday 29 October 2026, 9:30 local.
const NOW = new Date(2026, 9, 29, 9, 30);
const ymd = (d?: Date): string => (d ? `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}` : '-');

describe('date ranges', () => {
  it('today runs midnight to midnight, local', () => {
    const r = rangeOf({ preset: 'today' }, NOW);
    expect([ymd(r.from), ymd(r.to)]).toEqual(['2026-10-29', '2026-10-30']);
  });

  it('this week starts on Monday', () => {
    const r = rangeOf({ preset: 'thisWeek' }, NOW);
    expect([ymd(r.from), ymd(r.to)]).toEqual(['2026-10-26', '2026-11-2']);
  });

  it('past 7 and past 30 days include today', () => {
    expect(ymd(rangeOf({ preset: 'pastWeek' }, NOW).from)).toBe('2026-10-23');
    expect(ymd(rangeOf({ preset: 'last30' }, NOW).from)).toBe('2026-9-30');
  });

  it('this month, this year and a picked month', () => {
    expect(ymd(rangeOf({ preset: 'thisMonth' }, NOW).from)).toBe('2026-10-1');
    expect(ymd(rangeOf({ preset: 'thisYear' }, NOW).to)).toBe('2027-1-1');
    const feb = rangeOf({ preset: 'month', month: '2024-02' }, NOW);
    expect([ymd(feb.from), ymd(feb.to)]).toEqual(['2024-2-1', '2024-3-1']);
  });

  it('a date range includes both days, and survives being picked backwards', () => {
    const r = rangeOf({ preset: 'range', start: '2026-10-10', end: '2026-10-01' }, NOW);
    expect([ymd(r.from), ymd(r.to)]).toEqual(['2026-10-1', '2026-10-11']);
    expect(inRange(new Date(2026, 9, 10, 23, 59), r)).toBe(true);
    expect(inRange(new Date(2026, 9, 11, 0, 0), r)).toBe(false);
  });

  it('reads a picked day as local midnight, not UTC', () => {
    expect(parseDay('2026-09-01')?.getDate()).toBe(1);
    expect(parseDay('2026-02-31')).toBeUndefined();
  });

  it('describes what it covers', () => {
    expect(describeRange({ preset: 'month', month: '2026-03' }, NOW)).toBe('Mar 2026');
    expect(describeRange({ preset: 'today' }, NOW)).toBe('29 Oct 2026');
    expect(describeRange({ preset: 'range', start: '2026-10-01' }, NOW)).toBe('From 1 Oct 2026');
  });
});

const raw = (over: Record<string, unknown>): Record<string, unknown> => ({
  Id: 1, Title: 't', Ticket_x0020_ID: '261029PMW001', Status: 'New', Priority: 'Normal',
  Created: new Date(2026, 9, 29, 8).toISOString(), Modified: new Date(2026, 9, 29, 8).toISOString(),
  Assignedto0: { Id: 7, Title: 'Papathy', EMail: 'p@pmw' }, Author: { Id: 3, Title: 'Siti', EMail: 's@pmw' },
  Cc: [], ...over
});

describe('reading a ticket', () => {
  it('turns rich text into plain words', () => {
    expect(htmlToText('<div class="x"><p>Line one<br>two &amp; three&#58;</p><p>Next</p></div>')).toBe('Line one\ntwo & three:\nNext');
  });

  it('reads people, an empty Cc and a blank status', () => {
    const t = fromQueueItem(raw({ Status: null, Assignedto0: null, Cc: [{ Id: 9, Title: 'Boss', EMail: 'b@pmw' }] }));
    expect(t.status).toBe('New');
    expect(t.assignedTo).toBeUndefined();
    expect(t.cc.map((p) => p.name)).toEqual(['Boss']);
    expect(t.raisedBy?.email).toBe('s@pmw');
  });
});

describe('filtering', () => {
  const rows: QueueTicket[] = [
    fromQueueItem(raw({ Id: 1 })),
    fromQueueItem(raw({ Id: 2, Assignedto0: { Id: 8, Title: 'Ashraf', EMail: 'a@pmw' }, Title: 'SAP posting slow' })),
    fromQueueItem(raw({ Id: 3, Assignedto0: null, Status: 'Completed', Created: new Date(2026, 0, 5).toISOString() }))
  ];
  const all: QueueFilter = { assignees: [], date: { preset: 'all' }, statuses: [], search: '' };

  it('opens on the owner\'s own open tickets', () => {
    expect(applyFilter(rows, defaultFilter(7), NOW).map((t) => t.id)).toEqual([1]);
  });

  it('takes several assignees, and "unassigned"', () => {
    expect(applyFilter(rows, { ...all, assignees: [7, 8] }, NOW).map((t) => t.id)).toEqual([1, 2]);
    expect(applyFilter(rows, { ...all, assignees: [UNASSIGNED] }, NOW).map((t) => t.id)).toEqual([3]);
  });

  it('filters by the day it was raised', () => {
    expect(applyFilter(rows, { ...all, date: { preset: 'thisMonth' } }, NOW).map((t) => t.id)).toEqual([1, 2]);
  });

  it('needs every word searched for', () => {
    expect(applyFilter(rows, { ...all, search: 'sap ashraf' }, NOW).map((t) => t.id)).toEqual([2]);
    expect(applyFilter(rows, { ...all, search: 'sap papathy' }, NOW)).toEqual([]);
  });

  it('lists assignees with their counts', () => {
    expect(assigneesOf(rows).map((a) => `${a.person.name}:${a.count}`)).toEqual(['Ashraf:1', 'Papathy:1']);
  });

  it('sorts by priority, Critical first', () => {
    const p = ['Low', 'Critical', 'Normal'].map((pr, i) => fromQueueItem(raw({ Id: i, Priority: pr })));
    expect(sortTickets(p, 'priority', false).map((t) => t.priority)).toEqual(['Critical', 'Normal', 'Low']);
  });
});

describe('export', () => {
  const t = fromQueueItem(raw({ Title: 'Says "hi", then\nbreaks', Description: '<p>=HYPERLINK("x")</p>' }));

  it('offers every column by default, in sheet order', () => {
    expect(ALL_COLUMN_KEYS.length).toBe(EXPORT_COLUMNS.length);
    expect(chosenColumns(['status', 'number']).map((c) => c.key)).toEqual(['number', 'status']);
  });

  it('writes CSV Excel opens cleanly', () => {
    const csv = toCsv([t], chosenColumns(['number', 'title', 'description']));
    expect(csv.charCodeAt(0)).toBe(0xfeff);
    expect(csv).toContain('Ticket ID,Short description,Issue description\r\n');
    expect(csv).toContain('"Says ""hi"", then\nbreaks"');
    expect(csv).toContain(`"'=HYPERLINK(""x"")"`);
  });

  it('defuses formulas typed by staff', () => {
    expect(safeText('+60 12')).toBe("'+60 12");
    expect(safeText('normal')).toBe('normal');
  });

  it('names the file after what it holds', () => {
    expect(exportFileName(['Papathy', 'Oct 2026'], 'xlsx', NOW)).toBe('IT tickets - Papathy - Oct 2026 - 2026-10-29.xlsx');
    expect(exportFileName(['a/b:c'], 'csv', NOW)).toBe('IT tickets - abc - 2026-10-29.csv');
  });
});
