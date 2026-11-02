import { sortTickets } from './queueFilter';
import {
  addColumn, DEFAULT_COLUMNS, hiddenColumns, isDefault, moveColumn, parseTableColumns, removeColumn, TABLE_COLUMNS
} from './tableColumns';
import { fromQueueItem, QueueTicket } from './ticket';
import { durationText, hoursToComplete, timeToCompleteText } from './timeToComplete';

let nextId = 1;
const at = (d: number, h = 9): string => new Date(2026, 9, d, h).toISOString();
function t(over: Record<string, unknown>): QueueTicket {
  return fromQueueItem({ Id: nextId++, Title: 'x', Status: 'New', Created: at(1), Modified: at(1), Cc: [], ...over });
}

describe('time to complete', () => {
  it('reads days and hours, rounded to the hour', () => {
    expect(durationText(0.2)).toBe('Under 1 hr');
    expect(durationText(1)).toBe('1 hr');
    expect(durationText(5.4)).toBe('5 hrs');
    expect(durationText(24)).toBe('1 day');
    expect(durationText(49)).toBe('2 days 1 hr');
    expect(durationText(76.2)).toBe('3 days 4 hrs');
  });

  it('counts completed tickets only, from raised to completed (or the last update)', () => {
    expect(hoursToComplete(t({ Status: 'Completed', Created: at(1, 9), Completed_x0020_Date: at(3, 13) }))).toBe(52);
    expect(timeToCompleteText(t({ Status: 'Completed', Created: at(1, 9), Modified: at(1, 15) }))).toBe('6 hrs');
    expect(hoursToComplete(t({ Status: 'In progress' }))).toBeUndefined();
    expect(timeToCompleteText(t({ Status: 'Cancelled', Modified: at(5) }))).toBe('');
  });

  it('sorts quickest or slowest first, with unfinished tickets last either way', () => {
    const slow = t({ Status: 'Completed', Created: at(1), Completed_x0020_Date: at(9) });
    const quick = t({ Status: 'Completed', Created: at(1), Completed_x0020_Date: at(2) });
    const open = t({ Status: 'New' });
    expect(sortTickets([open, slow, quick], 'timeToComplete', false).map((x) => x.id)).toEqual([quick.id, slow.id, open.id]);
    expect(sortTickets([open, quick, slow], 'timeToComplete', true).map((x) => x.id)).toEqual([slow.id, quick.id, open.id]);
  });
});

describe('table columns', () => {
  it('starts with the queue as it was, plus time to complete', () => {
    expect(parseTableColumns(undefined)).toEqual([
      'number', 'title', 'category', 'priority', 'status', 'assignedTo', 'created', 'modified', 'timeToComplete'
    ]);
    expect(DEFAULT_COLUMNS.every((k) => TABLE_COLUMNS.some((c) => c.key === k))).toBe(true);
  });

  it('reads a saved choice back safely: Ticket ID first, unknown and repeated keys dropped', () => {
    expect(parseTableColumns('["status","gone","status","number","title"]')).toEqual(['number', 'status', 'title']);
    expect(parseTableColumns('not json')).toEqual(DEFAULT_COLUMNS);
    expect(parseTableColumns('[]')).toEqual(DEFAULT_COLUMNS);
    expect(parseTableColumns('{"a":1}')).toEqual(DEFAULT_COLUMNS);
  });

  it('moves, removes and adds, never shifting Ticket ID', () => {
    const keys = ['number', 'title', 'status', 'created'];
    expect(moveColumn(keys, 3, 1)).toEqual(['number', 'created', 'title', 'status']);
    expect(moveColumn(keys, 1, 0)).toEqual(['number', 'title', 'status', 'created']);
    expect(moveColumn(keys, 0, 2)).toEqual(keys);
    expect(removeColumn(keys, 'number')).toEqual(keys);
    expect(removeColumn(keys, 'status')).toEqual(['number', 'title', 'created']);
    expect(addColumn(keys, 'department')).toEqual(['number', 'title', 'status', 'created', 'department']);
    expect(addColumn(keys, 'title')).toEqual(keys);
    expect(hiddenColumns(keys).some((c) => c.key === 'title')).toBe(false);
    expect(isDefault(DEFAULT_COLUMNS.slice())).toBe(true);
    expect(isDefault(keys)).toBe(false);
  });
});
