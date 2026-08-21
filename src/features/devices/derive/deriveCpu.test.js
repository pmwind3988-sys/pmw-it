import { describe, it, expect } from 'vitest';
import { deriveCpu } from './deriveCpu.js';

describe('deriveCpu — generation', () => {
  it('uses the explicit "Nth Gen" prefix when the scan writes one', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuGeneration).toBe('13');
    expect(deriveCpu(['12th Gen Intel(R) Core(TM) i5-12400'], 'DDR4').cpuGeneration).toBe('12');
  });

  it('reads a 4-digit SKU beginning 10-14 as a two-digit generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i5-1035G1 CPU @ 1.00GHz'], 'DDR4').cpuGeneration)
      .toBe('10');
  });

  it('reads a 4-digit SKU beginning 2-9 as a one-digit generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i7-3667U CPU @ 2.00GHz'], 'DDR3').cpuGeneration)
      .toBe('3');
  });

  it('reads Core Ultra as a series, not an i-series generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) Ultra 5 125U'], 'DDR5').cpuGeneration).toBe('Ultra 1');
  });

  it('reads an AMD Ryzen series', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U with Radeon Graphics   '], 'DDR4').cpuGeneration)
      .toBe('Ryzen 7000');
  });

  it('has no generation for a Pentium', () => {
    expect(deriveCpu(['Intel(R) Pentium(R) Dual  CPU  E2160  @ 1.80GHz'], null).cpuGeneration)
      .toBe(null);
  });
});

describe('deriveCpu — vendor and model', () => {
  it('reads the vendor', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuVendor).toBe('Intel');
    expect(deriveCpu(['AMD Ryzen 5 7430U'], 'DDR4').cpuVendor).toBe('AMD');
  });

  it('trims the trailing whitespace the scan writes after AMD names', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U with Radeon Graphics         '], 'DDR4').cpuModel)
      .toBe('AMD Ryzen 5 7430U with Radeon Graphics');
  });
});

describe('deriveCpu — age band', () => {
  it('calls 10th generation and later Current', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuAgeBand).toBe('Current');
    expect(deriveCpu(['Intel(R) Core(TM) i5-1035G1'], 'DDR4').cpuAgeBand).toBe('Current');
  });

  it('calls Core Ultra Current', () => {
    expect(deriveCpu(['Intel(R) Core(TM) Ultra 5 125U'], 'DDR5').cpuAgeBand).toBe('Current');
  });

  it('calls 7th to 9th generation Aging', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i5-8250U'], 'DDR4').cpuAgeBand).toBe('Aging');
  });

  it('calls 6th generation and earlier Obsolete', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i7-3667U CPU @ 2.00GHz'], 'DDR3').cpuAgeBand)
      .toBe('Obsolete');
  });

  it('calls a Pentium with no generation Obsolete', () => {
    expect(deriveCpu(['Intel(R) Pentium(R) Dual  CPU  E2160  @ 1.80GHz'], null).cpuAgeBand)
      .toBe('Obsolete');
  });

  it('ranks AMD by series rather than inventing an Intel-comparable generation', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U'], 'DDR4').cpuAgeBand).toBe('Current');
    expect(deriveCpu(['AMD Ryzen 5 3500U'], 'DDR4').cpuAgeBand).toBe('Aging');
    expect(deriveCpu(['AMD Ryzen 3 2200U'], 'DDR4').cpuAgeBand).toBe('Obsolete');
  });

  it('calls DDR3 Obsolete even when no generation could be read', () => {
    expect(deriveCpu(['Some Unknown CPU'], 'DDR3').cpuAgeBand).toBe('Obsolete');
  });

  it('returns Unknown when there is nothing to go on', () => {
    expect(deriveCpu([], null).cpuAgeBand).toBe('Unknown');
  });
});
