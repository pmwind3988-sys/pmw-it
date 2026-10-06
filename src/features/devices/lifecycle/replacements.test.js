import { describe, it, expect } from 'vitest';
import { matchIncoming } from './matchIncoming.js';
import { replacementsFor, unanswered } from './replacements.js';

const scan = (over) => ({ computerName: 'CARMEN-HP2', serialNumber: 'NEW1', owner: 'Carmen', sourceFileName: 'c.txt', ...over });
const row = (over) => ({ id: 4, computerName: 'CARMEN-HP', serialNumber: 'OLD1', owner: 'Carmen', deviceType: 'Laptop', ...over });
const prompts = (incoming, existing) => replacementsFor(matchIncoming(incoming, existing), existing);

describe('replacementsFor', () => {
  it('asks when a person turns up on a new machine', () => {
    const [p] = prompts([scan()], [row()]);
    expect(p).toMatchObject({ key: 'c.txt|4', owner: 'Carmen', incomingName: 'CARMEN-HP2', old: { id: 4, computerName: 'CARMEN-HP' } });
  });

  it('compares names ignoring case and spacing', () => {
    expect(prompts([scan({ owner: ' carmen ' })], [row()])).toHaveLength(1);
  });

  it('asks once per old machine', () => {
    expect(prompts([scan()], [row(), row({ id: 5, computerName: 'CARMEN-PC' })])).toHaveLength(2);
  });

  it('does not ask about a machine already in the stash or retired', () => {
    expect(prompts([scan()], [row({ status: 'Spare' })])).toHaveLength(0);
  });

  it('does not ask again every time a known machine is re-scanned for the same person', () => {
    const existing = [row(), row({ id: 5, computerName: 'CARMEN-PC', serialNumber: 'PC1', deviceType: 'Desktop' })];
    expect(prompts([scan({ computerName: 'CARMEN-HP', serialNumber: 'OLD1' })], existing)).toHaveLength(0);
  });

  it('asks when a known machine is re-scanned under a new person who already has one', () => {
    const existing = [row({ owner: 'Farid' }), row({ id: 5, computerName: 'CARMEN-PC', serialNumber: 'PC1' })];
    expect(prompts([scan({ computerName: 'CARMEN-HP', serialNumber: 'OLD1' })], existing)).toHaveLength(1);
  });

  it('does not ask for a report with no owner', () => {
    expect(prompts([scan({ owner: null })], [row()])).toHaveLength(0);
  });
});

describe('unanswered', () => {
  it('counts prompts with no answer, skipping rows left out of the save', () => {
    const list = prompts([scan(), scan({ sourceFileName: 'd.txt', serialNumber: 'NEW2' })], [row()]);
    expect(unanswered(list, {})).toBe(2);
    expect(unanswered(list, { 'c.txt|4': 'keep' })).toBe(1);
    expect(unanswered(list, {}, new Set(['d.txt']))).toBe(1);
  });
});

describe('replacementsFor — same batch', () => {
  it('does not ask about an old machine whose own report is in the batch', () => {
    const existing = [row()];
    const incoming = [
      scan({ computerName: 'CARMEN-HP', serialNumber: 'OLD1', owner: 'Bob', sourceFileName: 'old.txt' }),
      scan({ owner: 'Carmen' }),
    ];
    expect(prompts(incoming, existing)).toHaveLength(0);
  });
});
