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

  it('reads an AMD Ryzen series and the architecture hidden inside it', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U with Radeon Graphics   '], 'DDR4').cpuGeneration)
      .toBe('Ryzen 7000 (Zen 3)');
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

  it('bands AMD on the same scale as Intel, by architecture rather than badge', () => {
    // A Zen 3 part is 11th-generation-era whatever its badge says, and a Zen
    // part is 7th — the same rung as the i5-8250U two tests up.
    expect(deriveCpu(['AMD Ryzen 5 7430U'], 'DDR4').cpuAgeBand).toBe('Current');
    expect(deriveCpu(['AMD Ryzen 5 3500U'], 'DDR4').cpuAgeBand).toBe('Aging');
    expect(deriveCpu(['AMD Ryzen 3 2200U'], 'DDR4').cpuAgeBand).toBe('Aging');
    expect(deriveCpu(['AMD Ryzen 5 2500U'], 'DDR3').cpuAgeBand).toBe('Obsolete');
  });

  it('calls DDR3 Obsolete even when no generation could be read', () => {
    expect(deriveCpu(['Some Unknown CPU'], 'DDR3').cpuAgeBand).toBe('Obsolete');
  });

  it('returns Unknown when there is nothing to go on', () => {
    expect(deriveCpu([], null).cpuAgeBand).toBe('Unknown');
  });
});

describe('deriveCpu — one scale for both vendors', () => {
  const rank = (model, ram = 'DDR4') => deriveCpu([model], ram).cpuGenerationRank;

  it('places an Intel part on its own generation', () => {
    expect(rank('13th Gen Intel(R) Core(TM) i7-1355U')).toBe(13);
    expect(rank('Intel(R) Core(TM) i5-8250U')).toBe(8);
  });

  it('continues the count through Core Ultra', () => {
    expect(rank('Intel(R) Core(TM) Ultra 5 125U', 'DDR5')).toBe(14);
    expect(rank('Intel(R) Core(TM) Ultra 7 265U', 'DDR5')).toBe(15);
  });

  it('reads the architecture digit in a 2022-or-later mobile Ryzen', () => {
    // Same 7000 badge, three years of architecture between them.
    expect(rank('AMD Ryzen 5 7530U with Radeon Graphics')).toBe(11);
    expect(rank('AMD Ryzen 7 7840U with Radeon 780M Graphics')).toBe(13);
    expect(rank('AMD Ryzen 3 7320U with Radeon Graphics')).toBe(10);
  });

  it('does not read a desktop model number as if it carried that digit', () => {
    // 7950X is Zen 4. Its third digit is a 5 and means nothing.
    expect(rank('AMD Ryzen 9 7950X 16-Core Processor')).toBe(13);
    expect(rank('AMD Ryzen 7 5800X 8-Core Processor')).toBe(11);
  });

  it('knows a mobile Ryzen runs a series behind its desktop namesake', () => {
    expect(rank('AMD Ryzen 5 3600 6-Core Processor')).toBe(10);
    expect(rank('AMD Ryzen 5 3500U with Radeon Vega Mobile Gfx')).toBe(8);
    expect(rank('AMD Ryzen 5 5600G with Radeon Graphics')).toBe(11);
  });

  it('reads a business Ryzen PRO part', () => {
    expect(rank('AMD Ryzen 5 PRO 5650U with Radeon Graphics')).toBe(11);
  });

  it('puts Ryzen AI on the newest rung', () => {
    expect(rank('AMD Ryzen AI 9 HX 370 w/ Radeon 890M', 'DDR5')).toBe(15);
    expect(deriveCpu(['AMD Ryzen AI 9 HX 370'], 'DDR5').cpuArchitecture).toBe('Zen 5');
  });

  it('makes the two vendors directly comparable', () => {
    const ryzen = deriveCpu(['AMD Ryzen 7 5825U with Radeon Graphics'], 'DDR4');
    const intel = deriveCpu(['11th Gen Intel(R) Core(TM) i5-1135G7'], 'DDR4');
    expect(ryzen.cpuGenerationRank).toBe(intel.cpuGenerationRank);
    expect(ryzen.cpuArchitecture).toBe('Zen 3');
  });

  it('has no rank for a processor it cannot place', () => {
    expect(rank('Intel(R) Pentium(R) Dual  CPU  E2160  @ 1.80GHz', null)).toBe(null);
  });
});

describe('deriveCpu — AMD parts that are not a plain Ryzen', () => {
  const cpu = (model, ram = 'DDR4') => deriveCpu([model], ram);

  it('reads a Threadripper, which carries no tier digit', () => {
    expect(cpu('AMD Ryzen Threadripper 3970X 32-Core Processor').cpuGenerationRank).toBe(10);
    expect(cpu('AMD Ryzen Threadripper PRO 5995WX 64-Cores').cpuGenerationRank).toBe(11);
  });

  it('places the Zen-based Athlons rather than giving up on them', () => {
    expect(cpu('AMD Athlon Gold 3150U with Radeon Graphics').cpuArchitecture).toBe('Zen');
    expect(cpu('AMD Athlon Silver 3050U with Radeon Graphics').cpuAgeBand).toBe('Aging');
    expect(cpu('AMD Athlon 3000G with Radeon Vega 3 Graphics').cpuGenerationRank).toBe(7);
  });

  it('calls everything AMD built before Zen Obsolete, DDR4 board or not', () => {
    // Without this they fall through to the RAM-type fallback, and a DDR4
    // board alone would call a 2014 APU Aging.
    expect(cpu('AMD A8-7410 APU with AMD Radeon R5 Graphics').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD A10-9600P RADEON R5').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD FX(tm)-8350 Eight-Core Processor').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD E1-6010 APU with AMD Radeon R2 Graphics').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD Athlon(tm) II X2 240 Processor').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD Athlon II X4 640 Processor').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD Phenom(tm) II X4 955 Processor').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD Sempron(tm) 145 Processor').cpuAgeBand).toBe('Obsolete');
    expect(cpu('AMD Turion(tm) II P540 Dual-Core Processor').cpuAgeBand).toBe('Obsolete');
  });

  it('does not read a pre-Zen Athlon as one of the Zen ones', () => {
    expect(cpu('AMD Athlon II X4 640 Processor').cpuGenerationRank).toBe(null);
    expect(cpu('AMD Athlon 64 X2 Dual Core 5600+').cpuAgeBand).toBe('Obsolete');
  });

  it('reads the vendor off the family name when the string omits "AMD"', () => {
    expect(cpu('Athlon(tm) II X2 240 Processor').cpuVendor).toBe('AMD');
    expect(cpu('FX(tm)-8350 Eight-Core Processor').cpuVendor).toBe('AMD');
    expect(cpu('Ryzen 5 7530U with Radeon Graphics').cpuVendor).toBe('AMD');
    expect(cpu('Intel(R) Core(TM) i5-8250U').cpuVendor).toBe('Intel');
  });
});
