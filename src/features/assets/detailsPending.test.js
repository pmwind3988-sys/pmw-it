import { describe, it, expect } from 'vitest';
import { needsDetails, missingDetails } from './detailsPending.js';
import { serialiseUnits } from './units.js';
import { TRACKED, BULK } from './assetKinds.js';

describe('needsDetails', () => {
  /**
   * A draft holds a boolean; a row read back out of SharePoint holds the word
   * the choice column stores. Both mean the same thing, and every caller would
   * otherwise have to know which kind of record it was holding.
   */
  it('reads the flag off a draft and off a saved row alike', () => {
    expect(needsDetails({ detailsPending: true })).toBe(true);
    expect(needsDetails({ detailsPending: 'Yes' })).toBe(true);
  });

  it('treats anything else as a finished row', () => {
    expect(needsDetails({ detailsPending: false })).toBe(false);
    expect(needsDetails({ detailsPending: 'No' })).toBe(false);
    expect(needsDetails({ detailsPending: null })).toBe(false);
    expect(needsDetails({})).toBe(false);
    expect(needsDetails(undefined)).toBe(false);
  });
});

describe('missingDetails', () => {
  it('names what is still blank, in words a person can go and find', () => {
    const missing = missingDetails({
      trackingMode: TRACKED, quantity: 1, doNumber: '', serialNumber: '', assetTag: '',
    });

    expect(missing).toEqual(['DO number', 'Serial number', 'Asset label', 'Photo']);
  });

  it('says nothing about what has already been filled in', () => {
    const missing = missingDetails({
      trackingMode: TRACKED,
      quantity: 1,
      doNumber: 'DO-8891',
      serialNumber: 'CN0MON001',
      assetTag: 'PMW-0142',
      photoUrl: 'https://contoso/photo.jpg',
    });

    expect(missing).toEqual([]);
  });

  /**
   * On a counted line the serials are on the individual items, so asking the
   * row for one would report ten monitors as missing a serial forever.
   */
  it('looks for a counted line\'s serials on its items', () => {
    const line = {
      trackingMode: BULK,
      quantity: 10,
      doNumber: 'DO-8891',
      assetTag: 'PMW-0142',
      photoUrl: 'https://contoso/photo.jpg',
      units: serialiseUnits([{ index: 0, serialNumber: 'CN0MON001' }]),
    };

    expect(missingDetails(line)).toEqual(['9 serial numbers']);
  });
});
