import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { importFiles, mergeImports, decodeReport } from './importFiles.js';
import { parseReport } from './parse/parseReport.js';

const fixture = (name) =>
  readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8');

/** Minimal stand-in for a browser File: name, lastModified, and the bytes the importer reads. */
const fakeFile = (name, lastModified, text = fixture(name)) => ({
  name,
  lastModified,
  text: async () => text,
  arrayBuffer: async () => new TextEncoder().encode(text).buffer,
});

describe('importFiles', () => {
  it('parses several files into records', async () => {
    const result = await importFiles([
      fakeFile('ASHRAF-PC_.txt', 1_760_000_000_000),
      fakeFile('PMWL034_.txt', 1_760_000_000_000),
    ]);
    expect(result.devices.map((d) => d.computerName)).toEqual(['ASHRAF-PC', 'PMWL034']);
    expect(result.rejected).toEqual([]);
  });

  it('rejects a file that is not a .txt', async () => {
    const result = await importFiles([fakeFile('report.pdf', 0, 'nonsense')]);
    expect(result.devices).toEqual([]);
    expect(result.rejected).toEqual([
      { fileName: 'report.pdf', reason: 'Not a .txt file' },
    ]);
  });

  it('rejects a .txt that is not a device report', async () => {
    const result = await importFiles([
      fakeFile('invoice.txt', 0, 'Dear team,\n\nPlease find the invoice attached.\n'),
    ]);
    expect(result.rejected).toEqual([
      { fileName: 'invoice.txt', reason: 'Not a device report — no known fields found' },
    ]);
  });

  it('keeps the newer of two files naming the same computer', async () => {
    const older = fakeFile('ASHRAF-PC_.txt', 1_700_000_000_000);
    const newer = fakeFile('[IT] ASHRAF-PC_.txt', 1_760_000_000_000, fixture('ASHRAF-PC_.txt'));
    const result = await importFiles([older, newer]);

    expect(result.devices).toHaveLength(1);
    expect(result.devices[0].sourceFileName).toBe('[IT] ASHRAF-PC_.txt');
    expect(result.rejected).toEqual([
      {
        fileName: 'ASHRAF-PC_.txt',
        reason: 'Duplicate of ASHRAF-PC — kept the newer scan from [IT] ASHRAF-PC_.txt',
      },
    ]);
  });

  it('reports a file it could not read without losing the rest of the batch', async () => {
    const broken = {
      name: 'broken.txt',
      lastModified: 0,
      text: async () => { throw new Error('disk gone'); },
      arrayBuffer: async () => { throw new Error('disk gone'); },
    };
    const result = await importFiles([broken, fakeFile('PMWL034_.txt', 0)]);

    expect(result.devices).toHaveLength(1);
    expect(result.rejected).toEqual([
      { fileName: 'broken.txt', reason: 'Could not read the file: disk gone' },
    ]);
  });
});

describe('mergeImports', () => {
  it('adds a second drop to the first instead of replacing it', async () => {
    const first = await importFiles([fakeFile('ASHRAF-PC_.txt', 1_760_000_000_000)]);
    const second = await importFiles([fakeFile('PMWL034_.txt', 1_760_000_000_000)]);

    const merged = mergeImports(first, second);
    expect(merged.devices.map((d) => d.computerName).sort()).toEqual(['ASHRAF-PC', 'PMWL034']);
  });

  it('keeps the newer scan when the same machine is dropped twice', async () => {
    const older = await importFiles([fakeFile('PMWL034_.txt', 1_760_000_000_000)]);
    const newer = await importFiles([fakeFile('PMWL034_.txt', 1_760_000_000_000)]);

    const merged = mergeImports(older, newer);
    expect(merged.devices).toHaveLength(1);
    expect(merged.rejected).toHaveLength(1);
    expect(merged.rejected[0].reason).toMatch(/Duplicate of PMWL034/);
  });

  it('carries the rejections of both drops', async () => {
    const first = await importFiles([fakeFile('report.pdf', 0, 'nonsense')]);
    const second = await importFiles([fakeFile('other.pdf', 0, 'nonsense')]);

    const merged = mergeImports(first, second);
    expect(merged.rejected.map((r) => r.fileName)).toEqual(['report.pdf', 'other.pdf']);
  });
});

describe('decodeReport', () => {
  const bytesOf = (text) => new TextEncoder().encode(text);

  /** UTF-16 code units of `text`, low byte first (little-endian) or high byte first. */
  const utf16 = (text, littleEndian) => {
    const out = [];
    for (const char of text) {
      const code = char.charCodeAt(0);
      const high = code >> 8;
      const low = code & 0xff;
      out.push(...(littleEndian ? [low, high] : [high, low]));
    }
    return new Uint8Array(out);
  };

  it('reads plain UTF-8', () => {
    expect(decodeReport(bytesOf('Computer Name:\nPC-1\n'))).toBe('Computer Name:\nPC-1\n');
  });

  it('drops a UTF-8 byte-order mark', () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...bytesOf('Name:\n')]);
    expect(decodeReport(bytes)).toBe('Name:\n');
  });

  it('reads UTF-16 little-endian behind its byte-order mark, without the mark in the text', () => {
    const text = 'Computer Name:\nAMIR-HP\n';
    const bytes = new Uint8Array([0xff, 0xfe, ...utf16(text, true)]);
    expect(decodeReport(bytes)).toBe(text);
  });

  it('reads UTF-16 big-endian behind its byte-order mark', () => {
    const text = 'Name:\n';
    const bytes = new Uint8Array([0xfe, 0xff, ...utf16(text, false)]);
    expect(decodeReport(bytes)).toBe(text);
  });

  it('reads the UTF-16 fixture as a device report, through the importer', async () => {
    const bytes = readFileSync(fileURLToPath(new URL('./__fixtures__/[ENGINEERING] AMIR-HP_.txt', import.meta.url)));
    const text = decodeReport(new Uint8Array(bytes));
    expect(parseReport(text).isReport).toBe(true);
    expect(parseReport(text).fields['Computer Name']).toEqual(['AMIR-HP']);

    const result = await importFiles([{
      name: '[ENGINEERING] AMIR-HP_.txt',
      lastModified: 0,
      arrayBuffer: async () => new Uint8Array(bytes).buffer,
    }]);
    expect(result.devices.map((d) => d.computerName)).toEqual(['AMIR-HP']);
  });
});
