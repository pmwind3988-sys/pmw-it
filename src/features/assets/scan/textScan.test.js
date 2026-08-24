import { describe, it, expect } from 'vitest';
import {
  newTextScan, recordReading, isSettled, settledValues, applyScannedFields, MAX_PASSES,
} from './textScan.js';

const reading = (fields) => ({
  serialNumber: '', partNumber: '', macAddress: '', assetTag: '',
  manufacturer: '', model: '', specSummary: '',
  additional: [], guessed: [], ...fields,
});

const readTwice = (fields) => recordReading(recordReading(newTextScan(), reading(fields)), reading(fields));

describe('recordReading — holding still until the text stops changing', () => {
  /**
   * The whole point of the settling rule. A single frame of a hand-held
   * camera misreads `8` as `B` often enough that accepting the first
   * answer would put a wrong serial number into the register.
   */
  it('does not accept a value read only once', () => {
    const scan = recordReading(newTextScan(), reading({ serialNumber: 'CN0ABC123456' }));

    expect(isSettled(scan, 'serialNumber')).toBe(false);
    expect(settledValues(scan).serialNumber).toBeUndefined();
  });

  it('accepts a value read the same way twice running', () => {
    const scan = readTwice({ serialNumber: 'CN0ABC123456' });

    expect(isSettled(scan, 'serialNumber')).toBe(true);
    expect(settledValues(scan).serialNumber).toBe('CN0ABC123456');
  });

  it('starts again when the reading changes', () => {
    let scan = recordReading(newTextScan(), reading({ serialNumber: 'CN0ABC123456' }));
    scan = recordReading(scan, reading({ serialNumber: 'CNOABC123456' }));

    expect(isSettled(scan, 'serialNumber')).toBe(false);
  });

  it('settles each field on its own', () => {
    let scan = recordReading(newTextScan(), reading({ serialNumber: 'CN0ABC123456', model: 'X' }));
    scan = recordReading(scan, reading({ serialNumber: 'CN0ABC123456', model: 'Y' }));

    expect(settledValues(scan)).toEqual({ serialNumber: 'CN0ABC123456' });
  });

  it('remembers that a settled value was a guess', () => {
    const scan = readTwice({ partNumber: 'LC-24B', guessed: ['partNumber'] });

    expect(scan.guessed).toContain('partNumber');
  });

  it('keeps every unplaced line it saw, without repeating one', () => {
    const scan = readTwice({ serialNumber: 'A1', additional: ['X9Y8', 'X9Y8'] });

    expect(scan.additional).toEqual(['X9Y8']);
  });

  it('gives up after a fixed number of passes rather than reading forever', () => {
    let scan = newTextScan();
    for (let pass = 0; pass < MAX_PASSES; pass += 1) {
      scan = recordReading(scan, reading({ serialNumber: `run-${pass}` }));
    }

    expect(scan.exhausted).toBe(true);
  });

  it('is not exhausted while it is still settling things', () => {
    expect(readTwice({ serialNumber: 'A1B2C3D4' }).exhausted).toBe(false);
  });
});

describe('applyScannedFields — a scan never overwrites what you typed', () => {
  const draft = (fields) => ({
    serialNumber: '', partNumber: '', specSummary: '', manufacturer: '',
    guessed: [], manualFields: [], additionalCodes: [], ...fields,
  });

  it('fills an empty field and marks a worked-out value as guessed', () => {
    const { record } = applyScannedFields(
      draft(),
      { serialNumber: 'CN0ABC123456', partNumber: 'LC-24B' },
      ['partNumber'],
    );

    expect(record.serialNumber).toBe('CN0ABC123456');
    expect(record.guessed).toEqual(['partNumber']);
  });

  it('leaves a hand-typed value alone and says what it held back', () => {
    const { record, heldBack } = applyScannedFields(
      draft({ serialNumber: 'TYPED-BY-HAND', manualFields: ['serialNumber'] }),
      { serialNumber: 'CN0ABC123456' },
    );

    expect(record.serialNumber).toBe('TYPED-BY-HAND');
    expect(heldBack).toEqual([{ field: 'serialNumber', value: 'CN0ABC123456' }]);
  });

  /**
   * A value an earlier scan guessed is not something anybody typed, so a
   * closer look at the label is allowed to correct it. That is the whole
   * reason for scanning the same box twice.
   */
  it('replaces a value an earlier scan guessed', () => {
    const { record, heldBack } = applyScannedFields(
      draft({ partNumber: 'WRONG', guessed: ['partNumber'] }),
      { partNumber: 'LC-24B' },
      ['partNumber'],
    );

    expect(record.partNumber).toBe('LC-24B');
    expect(heldBack).toEqual([]);
  });

  it('holds back a value that is already there but was never a guess', () => {
    const { record, heldBack } = applyScannedFields(
      draft({ specSummary: '8GB RAM' }),
      { specSummary: '16GB RAM' },
    );

    expect(record.specSummary).toBe('8GB RAM');
    expect(heldBack).toHaveLength(1);
  });

  it('ignores empty values instead of blanking the form with them', () => {
    const { record } = applyScannedFields(draft({ serialNumber: 'KEEP' }), { serialNumber: '' });

    expect(record.serialNumber).toBe('KEEP');
  });

  it('keeps unplaced lines against the row rather than dropping them', () => {
    const { record } = applyScannedFields(draft(), {}, [], ['X9Y8Z7']);

    expect(record.additionalCodes).toEqual(['X9Y8Z7']);
  });

  it('does not invent an additional-codes list on a record that has none', () => {
    const { record } = applyScannedFields({ serialNumber: '' }, { serialNumber: 'A1B2' }, [], ['X9']);

    expect(record.additionalCodes).toBeUndefined();
    expect(record.serialNumber).toBe('A1B2');
  });
});
