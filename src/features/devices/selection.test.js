import { describe, it, expect } from 'vitest';
import {
  isSelectable, selectableIds, toggleId, toggleAll, headerState,
  visibleSelection, selectedDevices, describeSelection,
} from './selection.js';

const rows = [
  { id: 1, computerName: 'A' },
  { id: 2, computerName: 'B' },
  { id: 3, computerName: 'C' },
];

const set = (...ids) => new Set(ids);

describe('isSelectable', () => {
  it('accepts a row that has a SharePoint id', () => {
    expect(isSelectable({ id: 4 })).toBe(true);
  });

  it('refuses a row with no id, because it cannot be removed either', () => {
    expect(isSelectable({ id: null })).toBe(false);
    expect(isSelectable({ id: undefined })).toBe(false);
    expect(isSelectable(null)).toBe(false);
  });

  it('accepts id 0, which is a value and not a blank', () => {
    expect(isSelectable({ id: 0 })).toBe(true);
  });
});

describe('selectableIds', () => {
  it('lists the ids in row order', () => {
    expect(selectableIds(rows)).toEqual([1, 2, 3]);
  });

  it('leaves out rows nothing can be done to', () => {
    expect(selectableIds([...rows, { id: null, computerName: 'D' }])).toEqual([1, 2, 3]);
  });
});

describe('toggleId', () => {
  it('ticks a row that was not ticked', () => {
    expect([...toggleId(set(1), 2)]).toEqual([1, 2]);
  });

  it('unticks a row that was', () => {
    expect([...toggleId(set(1, 2), 1)]).toEqual([2]);
  });

  it('does not mutate the set it was given', () => {
    const before = set(1);
    toggleId(before, 2);
    expect([...before]).toEqual([1]);
  });
});

describe('headerState', () => {
  it('is none when nothing is ticked', () => {
    expect(headerState(set(), rows)).toBe('none');
  });

  it('is some when part of the table is ticked', () => {
    expect(headerState(set(2), rows)).toBe('some');
  });

  it('is all when every row on screen is ticked', () => {
    expect(headerState(set(1, 2, 3), rows)).toBe('all');
  });

  it('is none for an empty table, so the header box is never half-ticked over nothing', () => {
    expect(headerState(set(1), [])).toBe('none');
  });

  it('ignores rows that cannot be ticked when deciding all', () => {
    // Otherwise a single id-less row would make "select all" unreachable.
    expect(headerState(set(1, 2, 3), [...rows, { id: null }])).toBe('all');
  });
});

describe('toggleAll', () => {
  it('ticks everything on screen when only some of it is ticked', () => {
    expect([...toggleAll(set(2), rows)]).toEqual([1, 2, 3]);
  });

  it('ticks everything on screen when none of it is', () => {
    expect([...toggleAll(set(), rows)]).toEqual([1, 2, 3]);
  });

  it('clears the lot when it is all ticked', () => {
    expect([...toggleAll(set(1, 2, 3), rows)]).toEqual([]);
  });
});

describe('visibleSelection', () => {
  it('drops a tick for a row the filters no longer show', () => {
    // The rule that matters: a search that narrows the table must not leave a
    // machine ticked off screen, or "Remove 3" would delete something nobody
    // can see.
    expect([...visibleSelection(set(1, 2, 3), [rows[0]])]).toEqual([1]);
  });

  it('clears the selection when the filters show nothing', () => {
    expect(visibleSelection(set(1, 2), []).size).toBe(0);
  });

  it('hands back the very same set when there was nothing to drop', () => {
    // Identity matters: a new Set on every render would re-run the effect
    // that prunes it, forever.
    const before = set(1, 2);
    expect(visibleSelection(before, rows)).toBe(before);
  });

  it('hands back the same empty set rather than a fresh one', () => {
    const before = set();
    expect(visibleSelection(before, rows)).toBe(before);
  });
});

describe('selectedDevices', () => {
  it('returns the ticked rows in the order the table shows them', () => {
    expect(selectedDevices(set(3, 1), rows).map((d) => d.computerName)).toEqual(['A', 'C']);
  });

  it('returns nothing when nothing is ticked', () => {
    expect(selectedDevices(set(), rows)).toEqual([]);
  });

  it('never returns a row that cannot be removed', () => {
    expect(selectedDevices(set(null), [{ id: null, computerName: 'D' }])).toEqual([]);
  });
});

describe('describeSelection', () => {
  const named = (...names) => names.map((computerName, id) => ({ id, computerName }));

  it('names one machine', () => {
    expect(describeSelection(named('PC1'))).toBe('PC1');
  });

  it('names a handful in full', () => {
    expect(describeSelection(named('A', 'B', 'C'))).toBe('A, B, C');
  });

  it('stops at four and counts the rest', () => {
    // The confirm sentence has to fit on one line. Twenty names would push the
    // button that removes them off the side of the bar.
    expect(describeSelection(named('A', 'B', 'C', 'D', 'E', 'F')))
      .toBe('A, B, C, D and 2 more');
  });

  it('says "1 more" rather than "1 more devices"', () => {
    expect(describeSelection(named('A', 'B', 'C', 'D', 'E'))).toBe('A, B, C, D and 1 more');
  });

  it('names a row whose computer name is missing rather than saying undefined', () => {
    expect(describeSelection([{ id: 1, computerName: null }])).toBe('an unnamed device');
  });

  it('is empty for an empty selection', () => {
    expect(describeSelection([])).toBe('');
  });
});
