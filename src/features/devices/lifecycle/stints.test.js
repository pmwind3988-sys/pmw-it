import { describe, it, expect } from 'vitest';
import {
  stintsOf, openStintOf, legacyStint, currentStint, timelineOf,
} from './stints.js';

const day = (d) => Date.UTC(2026, 0, d);
const machine = (over) => ({
  id: 4, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  status: null, createdOn: day(1), importedOn: day(2), ...over,
});
const stint = (over) => ({
  id: 1, deviceId: 4, owner: 'Farid', assignedOn: day(3), endedOn: day(10), endReason: 'Replaced', ...over,
});

describe('stintsOf / openStintOf', () => {
  it('keeps only this machine, newest first', () => {
    const list = [stint({ id: 1 }), stint({ id: 2, assignedOn: day(12), endedOn: null }), stint({ id: 3, deviceId: 9 })];
    expect(stintsOf(4, list).map((s) => s.id)).toEqual([2, 1]);
    expect(openStintOf(4, list).id).toBe(2);
    expect(openStintOf(9, [stint({ deviceId: 9 })])).toBeNull();
  });
});

describe('legacyStint', () => {
  it('is what the register implies about a machine with no history: since at least its creation', () => {
    expect(legacyStint(machine())).toMatchObject({
      id: null, deviceId: 4, owner: 'Amir', location: 'F1', department: 'ENGINEERING',
      assignedOn: day(1), assignedOnApprox: true, endedOn: null,
    });
  });

  it('falls back to the import date when there is no creation stamp', () => {
    expect(legacyStint(machine({ createdOn: null })).assignedOn).toBe(day(2));
  });

  it('is nothing for a machine with no owner', () => {
    expect(legacyStint(machine({ owner: null }))).toBeNull();
  });
});

describe('currentStint', () => {
  it('is the open stint when there is one', () => {
    expect(currentStint(machine(), [stint({ endedOn: null })]).id).toBe(1);
  });

  it('is the legacy stint for an in-use machine with no history at all', () => {
    expect(currentStint(machine(), []).id).toBeNull();
  });

  it('is nothing for a machine whose history has ended', () => {
    expect(currentStint(machine(), [stint()])).toBeNull();
  });

  it('is nothing for a spare machine', () => {
    expect(currentStint(machine({ status: 'Spare' }), [])).toBeNull();
  });
});

describe('timelineOf', () => {
  it('shows the time in the stash between two people', () => {
    const list = [
      stint({ id: 2, owner: 'Amir', assignedOn: day(20), endedOn: null, endReason: null }),
      stint({ id: 1, owner: 'Farid', assignedOn: day(3), endedOn: day(10), endReason: 'To stash' }),
    ];
    const entries = timelineOf(machine(), list);
    expect(entries.map((e) => e.kind)).toEqual(['stint', 'gap', 'stint']);
    expect(entries[0].current).toBe(true);
    expect(entries[1]).toEqual({ kind: 'gap', from: day(10), to: day(20), label: 'In IT Stash' });
  });

  it('ends with the stash or the graveyard for a machine nobody has now', () => {
    const entries = timelineOf(machine({ status: 'Retired', owner: null }), [stint({ endReason: 'Retired' })]);
    expect(entries[0]).toEqual({ kind: 'gap', from: day(10), to: null, label: 'In the Graveyard' });
  });

  it('shows the implied owner for a machine with no history yet', () => {
    const entries = timelineOf(machine(), []);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ kind: 'stint', owner: 'Amir', assignedOnApprox: true, current: true });
  });
});
