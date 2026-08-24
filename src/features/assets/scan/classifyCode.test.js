import { describe, it, expect } from 'vitest';
import { classifyCodes, serialScore, partScore } from './classifyCode.js';

const code = (rawValue, format = 'code_128') => ({ rawValue, format });
const again = (rawValue, format = 'code_128') => ({ rawValue, format, repeat: true });

describe('classifyCodes — explicit prefixes', () => {
  it('reads a prefixed serial and does not call it a guess', () => {
    const result = classifyCodes([code('S/N: CN0ABC123456')]);

    expect(result.serialNumber).toBe('CN0ABC123456');
    expect(result.guessed).not.toContain('serialNumber');
  });

  it.each([
    ['SN:5CG1234ABC', '5CG1234ABC'],
    ['Serial No: 5CG1234ABC', '5CG1234ABC'],
    ['SER 5CG1234ABC', '5CG1234ABC'],
  ])('accepts %s', (raw, expected) => {
    expect(classifyCodes([code(raw)]).serialNumber).toBe(expected);
  });

  it('reads a prefixed part number', () => {
    expect(classifyCodes([code('P/N: 5UF44AA')]).partNumber).toBe('5UF44AA');
  });

  it('reads a prefixed asset tag', () => {
    expect(classifyCodes([code('ASSET TAG: PMW-0142')]).assetTag).toBe('PMW-0142');
  });

  /**
   * The reason a separator is required. `SNK4820` is an ordinary serial, and
   * reading it as the prefix `SN` would file it as `K4820` — a confident wrong
   * answer that nothing on screen contradicts.
   */
  it('does not read a prefix without a separator', () => {
    expect(classifyCodes([code('SNK4820A1')]).serialNumber).toBe('SNK4820A1');
  });
});

describe('classifyCodes — MAC addresses', () => {
  it('recognises a separated MAC and normalises its case', () => {
    const result = classifyCodes([code('a4:bb:6d:1e:9f:02')]);
    expect(result.macAddress).toBe('A4:BB:6D:1E:9F:02');
    expect(result.serialNumber).toBe('');
  });

  it('accepts the dash-separated form', () => {
    expect(classifyCodes([code('A4-BB-6D-1E-9F-02')]).macAddress).toBe('A4-BB-6D-1E-9F-02');
  });

  /** Twelve bare hex characters is also a perfectly ordinary serial shape. */
  it('does not treat bare hex as a MAC', () => {
    const result = classifyCodes([code('A4BB6D1E9F02')]);
    expect(result.macAddress).toBe('');
    expect(result.serialNumber).toBe('A4BB6D1E9F02');
  });
});

describe('classifyCodes — retail barcodes', () => {
  /**
   * Every identical monitor on the pallet carries the same EAN. Reading one as
   * a serial is how twenty monitors collapse into one row.
   */
  it('files an EAN-13 as a part number, never a serial', () => {
    const result = classifyCodes([code('5901234123457', 'ean_13')]);

    expect(result.partNumber).toBe('5901234123457');
    expect(result.serialNumber).toBe('');
    expect(result.guessed).toContain('partNumber');
  });

  it('recognises a long digit run even when the format is unreported', () => {
    expect(classifyCodes([code('012345678905', '')]).partNumber).toBe('012345678905');
  });

  it('leaves the serial to the mixed code beside it', () => {
    const result = classifyCodes([
      code('5901234123457', 'ean_13'),
      code('CN0ABC123456'),
    ]);

    expect(result.partNumber).toBe('5901234123457');
    expect(result.serialNumber).toBe('CN0ABC123456');
  });
});

describe('classifyCodes — guessing from shape', () => {
  it('gives the serial to the longer mixed code', () => {
    const result = classifyCodes([code('5UF44AA'), code('CN0ABC1234567')]);

    expect(result.serialNumber).toBe('CN0ABC1234567');
    expect(result.partNumber).toBe('5UF44AA');
    expect(result.guessed).toEqual(expect.arrayContaining(['serialNumber', 'partNumber']));
  });

  it('keeps every code it cannot place', () => {
    const result = classifyCodes([
      code('CN0ABC1234567'),
      code('5UF44AA'),
      code('X1'),
      code('Y2'),
    ]);

    expect(result.additional).toContain('X1');
    expect(result.additional).toContain('Y2');
  });

  it('ignores blanks and repeats of the same code', () => {
    const result = classifyCodes([
      code('CN0ABC123456'),
      code('CN0ABC123456'),
      code('   '),
      code(''),
    ]);

    expect(result.serialNumber).toBe('CN0ABC123456');
    expect(result.partNumber).toBe('');
    expect(result.additional).toEqual([]);
  });

  it('handles a real HP label: serial, part number, MAC and a retail code', () => {
    const result = classifyCodes([
      code('5CG1234ABC'),
      code('P/N: 5UF44AA#ABU'),
      code('A4:BB:6D:1E:9F:02'),
      code('0195697123456', 'ean_13'),
    ]);

    expect(result.serialNumber).toBe('5CG1234ABC');
    expect(result.partNumber).toBe('5UF44AA#ABU');
    expect(result.macAddress).toBe('A4:BB:6D:1E:9F:02');
    expect(result.additional).toContain('0195697123456');
  });

  it('returns empty fields for no codes at all', () => {
    const result = classifyCodes([]);
    expect(result).toMatchObject({ serialNumber: '', partNumber: '', additional: [] });
  });

  it('is safe on undefined', () => {
    expect(classifyCodes(undefined).serialNumber).toBe('');
  });
});

describe('classifyCodes — part number versus serial number', () => {
  /**
   * The two-tabs case. The code that was already on the first box cannot be a
   * serial, whatever shape it happens to have, so the second tab keeps its own
   * serial instead of losing it to a coin toss between two look-alike codes.
   */
  it('files a code seen on an earlier box as the part number', () => {
    const result = classifyCodes([again('TAB10FE2024'), code('R52TC0ABCDE')]);

    expect(result.partNumber).toBe('TAB10FE2024');
    expect(result.serialNumber).toBe('R52TC0ABCDE');
  });

  it('does not let a repeated code take the serial even when it is alone', () => {
    const result = classifyCodes([again('TAB10FE2024')]);

    expect(result.partNumber).toBe('TAB10FE2024');
    expect(result.serialNumber).toBe('');
  });

  /** An explicit `S/N:` outranks the fact that the code has been seen before. */
  it('lets a prefix beat the repeat', () => {
    expect(classifyCodes([again('S/N: CN0ABC123')]).serialNumber).toBe('CN0ABC123');
  });

  /**
   * `#ABU` is HP's locale suffix and `/A` is Apple's. No serial scheme prints
   * either, so a box carrying only one of these has no serial to read — and
   * inventing one gives twenty identical items twenty identities.
   */
  it.each([
    ['5UF44AA#ABU'],
    ['MK2K3LL/A'],
  ])('leaves the serial empty for %s, which is a part number', (raw) => {
    const result = classifyCodes([code(raw)]);

    expect(result.partNumber).toBe(raw);
    expect(result.serialNumber).toBe('');
  });

  it('still reads the serial beside a vendor part number', () => {
    const result = classifyCodes([code('5UF44AA#ABU'), code('5CG1234ABC')]);

    expect(result.partNumber).toBe('5UF44AA#ABU');
    expect(result.serialNumber).toBe('5CG1234ABC');
  });

  it('reads a service tag as the serial, because that is what it is', () => {
    expect(classifyCodes([code('SVC TAG: 7XKL2Q3')]).serialNumber).toBe('7XKL2Q3');
  });

  it('reads MODEL as a part number, since every unit in the box shares it', () => {
    expect(classifyCodes([code('MODEL: SM-X210')]).partNumber).toBe('SM-X210');
  });

  /** Fifteen digits is one clear of the retail rule, and names one handset. */
  it('files an IMEI as the serial, not as a retail part number', () => {
    const result = classifyCodes([code('356938035643809')]);

    expect(result.serialNumber).toBe('356938035643809');
    expect(result.partNumber).toBe('');
  });

  it('keeps an IMEI verbatim when a serial was already read', () => {
    const result = classifyCodes([code('S/N: R52TC0ABCDE'), code('356938035643809')]);

    expect(result.serialNumber).toBe('R52TC0ABCDE');
    expect(result.partNumber).toBe('');
    expect(result.additional).toContain('356938035643809');
  });
});

describe('partScore', () => {
  it('reads vendor punctuation as the part number it is', () => {
    expect(partScore('5UF44AA#ABU')).toBeGreaterThan(serialScore('5UF44AA#ABU'));
    expect(partScore('MK2K3LL/A')).toBeGreaterThan(serialScore('MK2K3LL/A'));
  });

  it('does not out-argue an ordinary serial', () => {
    expect(partScore('CN0ABC123456')).toBeLessThan(serialScore('CN0ABC123456'));
  });

  it('treats letters with no digits as a model name', () => {
    expect(partScore('THINKPAD')).toBeGreaterThan(partScore('5CG1234ABC'));
  });
});

describe('serialScore', () => {
  it('prefers mixed letters and digits over pure digits', () => {
    expect(serialScore('CN0ABC123456')).toBeGreaterThan(serialScore('123456789012'));
  });

  it('penalises something too short to be unique', () => {
    expect(serialScore('AB1')).toBeLessThan(serialScore('AB123456'));
  });
});
