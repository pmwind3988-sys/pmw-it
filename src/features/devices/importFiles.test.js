import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { importFiles, mergeImports } from './importFiles.js';

const fixture = (name) =>
  readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8');

/** Minimal stand-in for a browser File: name, lastModified, text(). */
const fakeFile = (name, lastModified, text = fixture(name)) => ({
  name, lastModified, text: async () => text,
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
