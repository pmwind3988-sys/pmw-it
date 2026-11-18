import { describe, it, expect } from 'vitest';
import { newPicks, seePicks, takePick, visiblePicks } from './pickList';

const see = (picks, codes, now, skip) => seePicks(picks, codes, { now, skip });

describe('pick list', () => {
  it('one read is not shown', () => {
    const picks = see(newPicks(), ['ABC123'], 0);
    expect(visiblePicks(picks, 0)).toEqual([]);
  });

  it('two reads within 1.5s are shown', () => {
    let picks = see(newPicks(), [{ rawValue: 'abc123' }], 0);
    picks = see(picks, ['ABC123'], 300);
    expect(visiblePicks(picks, 300)).toEqual([{ code: 'ABC123', taken: false }]);
  });

  it('two reads 2s apart are not', () => {
    let picks = see(newPicks(), ['ABC123'], 0);
    picks = see(picks, ['ABC123'], 2000);
    expect(visiblePicks(picks, 2000)).toEqual([]);
  });

  it('a taken code reads taken and stays taken while in view', () => {
    let picks = see(see(newPicks(), ['ABC123'], 0), ['ABC123'], 100);
    picks = takePick(picks, 'abc123');
    picks = see(picks, ['ABC123'], 3000);
    picks = see(picks, ['ABC123'], 6000);
    expect(visiblePicks(picks, 6000)).toEqual([{ code: 'ABC123', taken: true }]);
  });

  it('a code unseen for 4s disappears, and a taken one can be shown untaken after', () => {
    let picks = see(see(newPicks(), ['ABC123'], 0), ['ABC123'], 100);
    picks = takePick(picks, 'ABC123');
    expect(visiblePicks(picks, 4000)).toHaveLength(1);
    expect(visiblePicks(picks, 4200)).toEqual([]);
    picks = see(see(picks, ['ABC123'], 9000), ['ABC123'], 9100);
    expect(visiblePicks(picks, 9100)).toEqual([{ code: 'ABC123', taken: false }]);
  });

  it('skip hides a code', () => {
    const skip = (code) => code === 'ONLIST1';
    let picks = see(newPicks(), ['ONLIST1', 'NEW1'], 0, skip);
    picks = see(picks, ['ONLIST1', 'NEW1'], 100, skip);
    expect(visiblePicks(picks, 100).map((pick) => pick.code)).toEqual(['NEW1']);
  });

  it('a tapped code stays as added even once the list holds it', () => {
    const listed = new Set();
    const skip = (code) => listed.has(code);
    let picks = see(see(newPicks(), ['SN1'], 0, skip), ['SN1'], 100, skip);
    picks = takePick(picks, 'SN1');
    listed.add('SN1');
    picks = see(picks, ['SN1'], 300, skip);
    expect(visiblePicks(picks, 300)).toEqual([{ code: 'SN1', taken: true }]);
  });

  it('at most six, oldest first', () => {
    let picks = newPicks();
    for (let i = 1; i <= 8; i += 1) {
      picks = see(see(picks, [`CODE${i}`], i * 10), [`CODE${i}`], i * 10 + 5);
    }
    expect(visiblePicks(picks, 100).map((pick) => pick.code))
      .toEqual(['CODE3', 'CODE4', 'CODE5', 'CODE6', 'CODE7', 'CODE8']);
  });
});
