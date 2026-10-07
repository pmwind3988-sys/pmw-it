import { describe, it, expect } from 'vitest';
import { TRACKED, BULK } from '../assetKinds.js';
import { newRun, addToRun, withoutSerial, undoLast, runSize, RUN_RESULT } from './serialRun.js';

const REGISTER = [
  { id: 1, trackingMode: BULK, units: JSON.stringify([{ index: 0, serialNumber: 'OLD1' }]) },
  { id: 2, trackingMode: TRACKED, serialNumber: 'LAPTOP1' },
];

describe('serial run', () => {
  it('counts one item per serial, plus the ones with none', () => {
    let { run } = addToRun(newRun(), 's1');
    ({ run } = addToRun(run, 'S2'));
    run = withoutSerial(run);
    expect(run.serials).toEqual(['S1', 'S2']);
    expect(runSize(run)).toBe(3);
    expect(runSize(undoLast(run))).toBe(2);
  });

  it('refuses a serial already in the run, on the receipt, or in the register', () => {
    const { run } = addToRun(newRun(), 'S1');
    expect(addToRun(run, 's1').result).toBe(RUN_RESULT.REPEAT);
    expect(addToRun(run, 'X9', { drafts: [{ units: JSON.stringify([{ index: 0, serialNumber: 'X9' }]) }] }).result).toBe(RUN_RESULT.ON_RECEIPT);
    expect(addToRun(run, 'old1', { assets: REGISTER }).result).toBe(RUN_RESULT.REGISTERED);
    expect(addToRun(run, 'LAPTOP1', { assets: REGISTER }).result).toBe(RUN_RESULT.REGISTERED);
  });
});
