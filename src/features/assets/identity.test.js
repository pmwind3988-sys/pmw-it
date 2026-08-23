import { describe, it, expect } from 'vitest';
import {
  assetKey, assetTitle, hasStableIdentity, indexByKey, indexByTag,
  normaliseCode, normaliseName,
} from './identity.js';
import { TRACKED, BULK } from './assetKinds.js';

describe('normalisation', () => {
  it('strips all spacing from a code, because printed spacing is not information', () => {
    expect(normaliseCode(' cn0abc 123 ')).toBe('CN0ABC123');
  });

  it('keeps the words apart in a name but collapses the runs', () => {
    expect(normaliseName('  ThinkPad   T14  Gen 3 ')).toBe('THINKPAD T14 GEN 3');
  });

  it('treats null as empty rather than as the string "null"', () => {
    expect(normaliseCode(null)).toBe('');
    expect(normaliseName(undefined)).toBe('');
  });
});

describe('assetKey — tracked items', () => {
  const laptop = {
    trackingMode: TRACKED,
    manufacturer: 'Dell',
    serialNumber: 'CN0ABC123',
    localId: 'local-1',
  };

  it('keys on manufacturer and serial', () => {
    expect(assetKey(laptop)).toBe('serial:DELL|CN0ABC123');
  });

  /** The dedupe that matters: the same machine scanned twice is one row. */
  it('gives the same key however the serial was typed', () => {
    expect(assetKey({ ...laptop, serialNumber: ' cn0abc 123 ' })).toBe(assetKey(laptop));
  });

  it('falls back to the sticker label when there is no serial', () => {
    expect(assetKey({ ...laptop, serialNumber: '', assetTag: 'PMW-0142' }))
      .toBe('tag:PMW-0142');
  });

  it('falls back to a local id when there is neither, and says so', () => {
    const key = assetKey({ ...laptop, serialNumber: '', assetTag: '' });

    expect(key).toBe('local:local-1');
    expect(hasStableIdentity(key)).toBe(false);
  });

  it('counts a serial key as stable', () => {
    expect(hasStableIdentity(assetKey(laptop))).toBe(true);
  });
});

describe('assetKey — bulk lines', () => {
  const mouse = {
    trackingMode: BULK,
    category: 'Mouse',
    manufacturer: 'Logitech',
    model: 'B100',
  };

  it('keys on what it is, not which one', () => {
    expect(assetKey(mouse)).toBe('bulk:MOUSE|LOGITECH|B100');
  });

  /** A second bag of the same mice must find the first bag's row. */
  it('matches another bag of the same thing', () => {
    expect(assetKey({ ...mouse, manufacturer: 'logitech ' })).toBe(assetKey(mouse));
  });

  it('ignores a serial that happened to be scanned onto a bulk row', () => {
    expect(assetKey({ ...mouse, serialNumber: 'ABC123' })).toBe(assetKey(mouse));
  });
});

describe('indexes', () => {
  const rows = [
    { assetKey: 'serial:DELL|A1', assetTag: 'PMW-1', id: 1 },
    { assetKey: 'bulk:MOUSE||B100', assetTag: '', id: 2 },
  ];

  it('indexes by key', () => {
    expect(indexByKey(rows).get('serial:DELL|A1').id).toBe(1);
  });

  it('indexes tags normalised, so a stray space still collides', () => {
    expect(indexByTag(rows).get('PMW-1').id).toBe(1);
    expect(indexByTag([{ assetTag: ' pmw-1 ' }]).has('PMW-1')).toBe(true);
  });

  it('leaves rows with no tag out of the tag index', () => {
    expect(indexByTag(rows).size).toBe(1);
  });
});

describe('assetTitle', () => {
  it('reads as make, model and serial', () => {
    expect(assetTitle({ manufacturer: 'Dell', model: 'P2422H', serialNumber: 'CN0ABC' }))
      .toBe('Dell P2422H — CN0ABC');
  });

  it('drops the dash when there is no serial', () => {
    expect(assetTitle({ manufacturer: 'Logitech', model: 'B100' })).toBe('Logitech B100');
  });

  it('falls back to the category rather than going blank', () => {
    expect(assetTitle({ category: 'Cable' })).toBe('Cable');
  });
});
