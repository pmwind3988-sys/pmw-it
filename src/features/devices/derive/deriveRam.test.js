import { describe, it, expect } from 'vitest';
import { deriveRam } from './deriveRam.js';

const twoByFour = [
  '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
  '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
  'Total Slots: 2 | Used Slots: 2',
];

describe('deriveRam', () => {
  it('sums the sticks for the installed figure', () => {
    expect(deriveRam(twoByFour, ['8 GB']).installedRamGB).toBe(8);
  });

  it('flags the iGPU reservation gap: 15 GB reported is a 16 GB machine', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | Micron Technology | 4ATF1G64HZ-3G2F1',
        '8 GB | DDR4 | 3200 MHz | Micron Technology | 4ATF1G64HZ-3G2F1',
        'Total Slots: 2 | Used Slots: 2'],
      ['15 GB'],
    );
    expect(result.installedRamGB).toBe(16);
    expect(result.reportedRamGB).toBe(15);
    expect(result.ramDiscrepancy).toBe(true);
  });

  it('does not flag a discrepancy when the two agree', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramDiscrepancy).toBe(false);
  });

  it('reports a free slot as upgradable', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | Kingston | HP32D4S2S8MR-8', 'Total Slots: 2 | Used Slots: '],
      ['8 GB'],
    );
    expect(result.ramSlotsUsed).toBe(1);
    expect(result.ramSlotsTotal).toBe(2);
    expect(result.ramUpgradable).toBe(true);
  });

  it('is not upgradable when every slot is filled', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramUpgradable).toBe(false);
  });

  it('takes the slowest speed when sticks differ', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | A | 1', '8 GB | DDR4 | 2667 MHz | B | 2'],
      ['16 GB'],
    );
    expect(result.ramSpeedMhz).toBe(2667);
  });

  it('reports the most common stick type', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramType).toBe('DDR4');
  });

  it('returns Unknown type when the scan could not read it', () => {
    const result = deriveRam(
      ['2 GB | Unknown | 333 MHz | Manufacturer1 | PartNum1', 'Total Slots: 2 | Used Slots: '],
      ['2 GB'],
    );
    expect(result.ramType).toBe('Unknown');
    expect(result.installedRamGB).toBe(2);
  });

  it('handles a report with no RAM block at all', () => {
    const result = deriveRam([], []);
    expect(result.installedRamGB).toBe(null);
    expect(result.reportedRamGB).toBe(null);
    expect(result.ramDiscrepancy).toBe(false);
  });
});
