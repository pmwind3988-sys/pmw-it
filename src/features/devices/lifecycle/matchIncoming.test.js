import { describe, it, expect } from 'vitest';
import {
  MATCH, matchIncoming, noticeFor, NEW_MACHINE,
} from './matchIncoming.js';

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

  it('does not claim to bring a machine back when the report is older than its move', () => {
    const [m] = matchIncoming([scan({ scannedOn: 1000 })], [row({ status: 'Spare', statusChangedOn: 5000 })]);
    expect(noticeFor(m)).toMatch(/Older than its move to IT Stash: specs updated, status kept/);
    const [r] = matchIncoming([scan({ scannedOn: 1000 })], [row({ status: 'Retired', statusChangedOn: 5000 })]);
    expect(noticeFor(r)).toMatch(/Older than its move to the Graveyard/);
  });

  it('says nothing for an ordinary re-scan', () => {
    expect(noticeFor(matchIncoming([scan()], [row()])[0])).toBeNull();
  });
});

describe('matchIncoming with choices', () => {
  it('saves a report as a new machine when told to, and says what it leaves alone', () => {
    const [m] = matchIncoming([scan()], [row()], { 'a.txt': NEW_MACHINE });
    expect(m.kind).toBe(MATCH.NEW);
    expect(m.existing).toBeNull();
    expect(m.auto.id).toBe(4);
    expect(m.note).toMatch(/new machine, as chosen\. AMIR-HP is left as it is\. Both carry serial/);
  });

  it('updates the machine chosen instead of the one it matched', () => {
    const other = row({ id: 9, computerName: 'SPARE-01', serialNumber: 'XYZ' });
    const [m] = matchIncoming([scan()], [row(), other], { 'a.txt': 9 });
    expect(m.existing.id).toBe(9);
    expect(m.kind).toBe(MATCH.RENAMED);
    expect(m.note).toMatch(/Updates SPARE-01, as chosen -- it takes the name AMIR-HP/);
  });

  it('updates a machine it found nothing for', () => {
    const [m] = matchIncoming([scan({ serialNumber: null, computerName: 'NEW-PC' })], [row()], { 'a.txt': 4 });
    expect(m.existing.id).toBe(4);
    expect(m.chosen).toBe(true);
  });

  it('falls back to the automatic match when the chosen machine is gone', () => {
    const [m] = matchIncoming([scan()], [row()], { 'a.txt': 99 });
    expect(m.existing.id).toBe(4);
    expect(m.chosen).toBeUndefined();
  });

  it('lets one report update a machine, and skips a second that would update it too', () => {
    const other = scan({ serialNumber: null, computerName: 'OTHER', sourceFileName: 'b.txt' });
    const matches = matchIncoming([scan(), other], [row()], { 'b.txt': 4 });
    expect(matches[1].existing.id).toBe(4);
    expect(matches[0].kind).toBe(MATCH.DUPLICATE_TARGET);
    expect(matches[0].note).toMatch(/AMIR-HP is already updated by OTHER/);
  });
});
