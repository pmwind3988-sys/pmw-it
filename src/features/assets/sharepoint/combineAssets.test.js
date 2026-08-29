import { describe, it, expect } from 'vitest';
import { repointFor } from './combineAssets.js';

const kept = { id: 4, assetKey: 'bulk:MONITOR|DELL|P2422H', title: 'Dell P2422H' };

const handover = (extra = {}) => ({
  id: 55,
  assetId: 9,
  assetKey: 'serial:DELL|SN-2',
  itemTitle: 'Dell — SN-2',
  unitIndex: null,
  ...extra,
});

describe('repointFor', () => {
  it('moves a handover onto the line its row became', () => {
    const patch = repointFor({ was: handover(), row: {}, unitIndex: 1 }, kept);

    expect(patch).toEqual({
      assetKey: 'bulk:MONITOR|DELL|P2422H',
      assetId: 4,
      unitIndex: 1,
    });
  });

  it('leaves the title alone — it says what went out that day', () => {
    const patch = repointFor({ was: handover(), row: {}, unitIndex: 1 }, kept);

    expect('itemTitle' in patch).toBe(false);
  });

  it('carries a folded row serial onto a handover that had none', () => {
    const patch = repointFor(
      { was: handover(), row: { serialNumber: 'SN-2' }, unitIndex: 1 },
      kept,
    );

    // Without this nothing would say which of ten identical monitors went out:
    // the row title used to carry it, and the line has one title for all ten.
    expect(patch.serialNumber).toBe('SN-2');
  });

  it('never writes over a serial the handover already names', () => {
    const patch = repointFor(
      { was: handover({ serialNumber: 'SN-EXACT' }), row: { serialNumber: 'SN-ROW' }, unitIndex: 1 },
      kept,
    );

    expect('serialNumber' in patch).toBe(false);
  });

  it('sends nothing for a handover that already says all of it', () => {
    const already = handover({
      assetId: kept.id, assetKey: kept.assetKey, unitIndex: 1, serialNumber: 'SN-2',
    });

    expect(repointFor({ was: already, row: { serialNumber: 'SN-2' }, unitIndex: 1 }, kept))
      .toBeNull();
  });

  it('still moves the keeper own handovers when only its key changed', () => {
    const onKeeper = handover({ assetId: kept.id, unitIndex: 0, serialNumber: 'SN-1' });

    expect(repointFor({ was: onKeeper, row: { serialNumber: 'SN-1' }, unitIndex: 0 }, kept))
      .toEqual({ assetKey: kept.assetKey });
  });
});
