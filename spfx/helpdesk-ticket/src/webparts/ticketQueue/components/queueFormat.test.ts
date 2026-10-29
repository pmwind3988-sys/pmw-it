import { assigneeSummary, assigneeNames, exportLines, formatWhen, pageCount, parseColumns, parseFilter, sortDescription } from './queueFormat';
import { ALL_COLUMN_KEYS } from '../queue/exportColumns';
import { defaultFilter, UNASSIGNED } from '../queue/queueFilter';
import { PersonRef } from '../queue/ticket';

const me: PersonRef = { id: 7, name: 'Papathy, Ashraf', email: 'ashraf@example.com' };
const names: Record<number, string> = { 3: 'Lim, Wei', 9: 'Aziz, Nur' };
const nameOf = (id: number): string => names[id] ?? 'Someone else';

describe('formatWhen', () => {
  it('reads a local instant as day, short month, year and a 12-hour time', () => {
    expect(formatWhen(new Date(2026, 9, 29, 14, 5).toISOString())).toBe('29 Oct 2026, 2:05 PM');
  });

  it('shows midnight and noon as 12', () => {
    expect(formatWhen(new Date(2026, 9, 29, 0, 7).toISOString())).toBe('29 Oct 2026, 12:07 AM');
    expect(formatWhen(new Date(2026, 9, 29, 12, 0).toISOString())).toBe('29 Oct 2026, 12:00 PM');
  });

  it('is empty for a value that is not a date', () => {
    expect(formatWhen('')).toBe('');
    expect(formatWhen('not a date')).toBe('');
  });
});

describe('assignee words', () => {
  it('says Everyone for no choice, and Me for only me', () => {
    expect(assigneeSummary([], me, nameOf)).toBe('Everyone');
    expect(assigneeSummary([me.id], me, nameOf)).toBe('Me');
  });

  it('counts the others after the first name', () => {
    expect(assigneeSummary([me.id, 3, 9], me, nameOf)).toBe('Me + 2 others');
    expect(assigneeSummary([me.id, UNASSIGNED], me, nameOf)).toBe('Me + 1 other');
  });

  it('names one other person in full, and Unassigned as itself', () => {
    expect(assigneeSummary([3], me, nameOf)).toBe('Lim, Wei');
    expect(assigneeSummary([UNASSIGNED], me, nameOf)).toBe('Unassigned');
  });

  it('lists me first, Unassigned next, then the rest A–Z', () => {
    expect(assigneeNames([9, UNASSIGNED, 3, me.id], me, nameOf, 'Me')).toEqual([
      'Me',
      'Unassigned',
      'Aziz, Nur',
      'Lim, Wei'
    ]);
  });
});

describe('saved filter', () => {
  it('falls back to the default when nothing or bad JSON is stored', () => {
    expect(parseFilter(undefined, me.id)).toEqual(defaultFilter(me.id));
    expect(parseFilter('{not json', me.id)).toEqual(defaultFilter(me.id));
    expect(parseFilter('[1,2]', me.id)).toEqual(defaultFilter(me.id));
  });

  it('keeps good fields and drops the rest one by one', () => {
    const raw = JSON.stringify({
      assignees: [3, 'x', 9],
      date: { preset: 'nonsense', month: '2026-10' },
      statuses: ['Completed', 'Bogus', 'New'],
      search: 'sap'
    });
    const f = parseFilter(raw, me.id);
    expect(f.assignees).toEqual([3, 9]);
    expect(f.date).toEqual({ preset: 'all', month: '2026-10', start: undefined, end: undefined });
    expect(f.statuses).toEqual(['New', 'Completed']);
    expect(f.search).toBe('sap');
  });
});

describe('saved export columns', () => {
  it('means all columns when nothing usable is stored', () => {
    expect(parseColumns(undefined)).toEqual(ALL_COLUMN_KEYS);
    expect(parseColumns('[]')).toEqual(ALL_COLUMN_KEYS);
    expect(parseColumns('["nope"]')).toEqual(ALL_COLUMN_KEYS);
    expect(parseColumns('{}')).toEqual(ALL_COLUMN_KEYS);
  });

  it('drops unknown keys and keeps the sheet order', () => {
    expect(parseColumns('["status","number","bogus"]')).toEqual(['number', 'status']);
  });
});

describe('paging and export text', () => {
  it('counts pages of fifty, never fewer than one', () => {
    expect(pageCount(0)).toBe(1);
    expect(pageCount(50)).toBe(1);
    expect(pageCount(51)).toBe(2);
    expect(pageCount(230)).toBe(5);
  });

  it('describes the sort in words for the sheet', () => {
    expect(sortDescription('created', true)).toBe('Raised on (newest first)');
    expect(sortDescription('priority', false)).toBe('Priority (Critical first)');
  });

  it('writes the about-sheet lines in order, with blanks shown as a dash', () => {
    const lines = exportLines(
      {
        exportedBy: 'Papathy, Ashraf',
        assignees: 'Me',
        raised: 'Any time',
        statuses: [],
        search: '  ',
        sortText: 'Raised on (newest first)',
        fileParts: []
      },
      12,
      new Date(2026, 9, 29, 9, 30)
    );
    expect(lines[0]).toEqual(['Exported by', 'Papathy, Ashraf']);
    expect(lines[1]).toEqual(['Exported on', '2026-10-29 09:30']);
    expect(lines[4]).toEqual(['Status', 'Any']);
    expect(lines[5]).toEqual(['Search', '—']);
    expect(lines[7]).toEqual(['Tickets', '12']);
  });
});
