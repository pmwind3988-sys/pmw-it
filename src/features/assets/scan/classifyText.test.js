import { describe, it, expect } from 'vitest';
import { cleanLines, isSpecLine, readTextFields } from './classifyText.js';

const lines = (...values) => values.map((text) => ({ text, confidence: 90 }));

describe('cleanLines', () => {
  it('collapses the whitespace an uneven label produces', () => {
    expect(cleanLines(lines('S/N:   5CG1234ABC  '))).toEqual(['S/N: 5CG1234ABC']);
  });

  it('drops a line with nothing readable in it', () => {
    expect(cleanLines(lines('|| ~~ ||', '...'))).toEqual([]);
  });

  it('drops a single stray character', () => {
    expect(cleanLines(lines('l'))).toEqual([]);
  });

  /**
   * Recognition returns its own confidence, and a low one on a label is
   * almost always the printed noise around the text rather than the text.
   * Keeping it would let a misread character reach a serial number field.
   */
  it('drops a line the reader was not confident about', () => {
    expect(cleanLines([{ text: '5CG1234ABC', confidence: 20 }])).toEqual([]);
  });

  it('keeps the same line only once however often it is read', () => {
    expect(cleanLines(lines('LATITUDE 5540', 'LATITUDE 5540'))).toEqual(['LATITUDE 5540']);
  });
});

describe('isSpecLine', () => {
  it.each([
    '16GB RAM',
    '512GB SSD',
    'Intel Core i7-1355U',
    'AMD Ryzen 7 5825U',
    '2.4 GHz',
    '1TB NVMe',
    '24" FHD 1920x1080',
  ])('recognises %s as a specification', (text) => {
    expect(isSpecLine(text)).toBe(true);
  });

  it.each([
    'CN0ABC123456',
    '5UF44AA#ABU',
    'PMW-0142',
    'Latitude 5540',
  ])('does not mistake %s for a specification', (text) => {
    expect(isSpecLine(text)).toBe(false);
  });
});

describe('readTextFields — what each line turns out to be', () => {
  it('files a prefixed serial and a prefixed part number separately', () => {
    const result = readTextFields(lines('S/N: CN0ABC123456', 'P/N: 5UF44AA#ABU'));

    expect(result.serialNumber).toBe('CN0ABC123456');
    expect(result.partNumber).toBe('5UF44AA#ABU');
    // Read outright, not worked out — the label said which was which.
    expect(result.guessed).toEqual([]);
  });

  it('collects the specification lines into one summary', () => {
    const result = readTextFields(lines('16GB RAM', '512GB SSD', 'S/N: CN0ABC123456'));

    expect(result.specSummary).toBe('16GB RAM, 512GB SSD');
    expect(result.guessed).toContain('specSummary');
    expect(result.serialNumber).toBe('CN0ABC123456');
  });

  it('takes a make it recognises off a line of its own', () => {
    const result = readTextFields(lines('DELL', 'S/N: CN0ABC123456'));

    expect(result.manufacturer).toBe('Dell');
    expect(result.guessed).toContain('manufacturer');
  });

  it('reads a labelled make and model without guessing', () => {
    const result = readTextFields(lines('Brand: Lenovo', 'Model: ThinkPad T14 Gen 4'));

    expect(result.manufacturer).toBe('Lenovo');
    expect(result.model).toBe('ThinkPad T14 Gen 4');
    expect(result.guessed).toEqual([]);
  });

  /**
   * A labelled `Model:` carrying a part-number shape is the part number —
   * which is what the barcode classifier already decided for the same
   * prefix, and the two must not disagree.
   */
  it('files a part-shaped model line as the part number', () => {
    const result = readTextFields(lines('Model: 5UF44AA#ABU'));

    expect(result.partNumber).toBe('5UF44AA#ABU');
    expect(result.model).toBe('');
  });

  it('falls back to shape when no line is labelled', () => {
    const result = readTextFields(lines('CN0ABC123456', 'LC-24B'));

    expect(result.serialNumber).toBe('CN0ABC123456');
    expect(result.partNumber).toBe('LC-24B');
    expect(result.guessed).toEqual(expect.arrayContaining(['serialNumber', 'partNumber']));
  });

  it('recognises a MAC address printed on the label', () => {
    expect(readTextFields(lines('MAC A4:BB:6D:1E:9F:02')).macAddress).toBe('A4:BB:6D:1E:9F:02');
  });

  it('keeps what it could not place rather than dropping it', () => {
    const result = readTextFields(lines(
      'S/N: CN0ABC123456',
      'P/N: 5UF44AA',
      'ASSET: PMW-0142',
      'MAC A4:BB:6D:1E:9F:02',
      'X9Y8Z7W6V5',
    ));

    expect(result.additional).toContain('X9Y8Z7W6V5');
  });

  it('returns empty fields for a frame with nothing on it', () => {
    const result = readTextFields([]);

    expect(result.serialNumber).toBe('');
    expect(result.specSummary).toBe('');
    expect(result.guessed).toEqual([]);
  });
});
