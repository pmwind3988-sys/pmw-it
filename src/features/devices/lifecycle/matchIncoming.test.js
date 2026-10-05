import { describe, it, expect } from 'vitest';
import { MATCH, matchIncoming, noticeFor } from './matchIncoming.js';

const scan = (over) => ({ computerName: 'AMIR-HP', serialNumber: '5CG8241KQZ', owner: 'Amir', sourceFileName: 'a.txt', ...over });
const row = (over) => ({ id: 4, computerName: 'AMIR-HP', serialNumber: '5CG8241KQZ', owner: 'Amir', ...over });

describe('matchIncoming', () => {
  it('finds a machine by its serial', () => {
    const [m] = matchIncoming([scan()], [row()]);
    expect(m.kind).toBe(MATCH.SAME);
    expect(m.existing.id).toBe(4);
  });

  it('sees a rename: same serial, new name', () => {
    const [m] = matchIncoming([scan({ computerName: 'AISYAH-HP' })], [row()]);
    expect(m.kind).toBe(MATCH.RENAMED);
    expect(m.note).toMatch(/Renamed from AMIR-HP/);
  });

  it('matches serials whatever their spacing or case', () => {
    expect(matchIncoming([scan({ serialNumber: '5cg8241kqz' })], [row({ serialNumber: '5CG 8241 KQZ' })])[0].kind).toBe(MATCH.SAME);
  });

  it('falls back to the name for a report with no serial', () => {
    expect(matchIncoming([scan({ serialNumber: null })], [row()])[0].kind).toBe(MATCH.SAME);
  });

  it('writes the serial onto a row that had none', () => {
    const [m] = matchIncoming([scan()], [row({ serialNumber: null })]);
    expect(m.kind).toBe(MATCH.SAME);
    expect(m.existing.id).toBe(4);
  });

  it('treats a reused name with a different serial as a new machine', () => {
    const [m] = matchIncoming([scan({ serialNumber: 'NEWSERIAL1' })], [row()]);
    expect(m.kind).toBe(MATCH.NAME_REUSED);
    expect(m.existing).toBeNull();
    expect(m.note).toMatch(/different machine \(serial 5CG8241KQZ\)/);
  });

  it('flags two reports in one drop claiming one serial', () => {
    const matches = matchIncoming([scan(), scan({ computerName: 'OTHER', sourceFileName: 'b.txt' })], []);
    expect(matches.map((m) => m.kind)).toEqual([MATCH.NEW, MATCH.DUPLICATE_SERIAL]);
  });

  it('is new when nothing matches', () => {
    expect(matchIncoming([scan()], [])[0].kind).toBe(MATCH.NEW);
  });
});

describe('noticeFor', () => {
  it('says a retired machine scanned again is being brought back', () => {
    const [m] = matchIncoming([scan()], [row({ status: 'Retired' })]);
    expect(noticeFor(m)).toMatch(/Brought back from the Graveyard/);
  });

  it('says nothing for an ordinary re-scan', () => {
    expect(noticeFor(matchIncoming([scan()], [row()])[0])).toBeNull();
  });
});
