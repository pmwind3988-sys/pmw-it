import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from '../parse/parseReport.js';
import { deriveScanInfo, scanRank } from './deriveScanInfo.js';

const read = (path) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8');
const fieldsOf = (text) => parseReport(text).fields;

describe('deriveScanInfo', () => {
  it('reads a full version 2 scan, run without administrator rights', () => {
    expect(deriveScanInfo(fieldsOf(read('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt'))))
      .toEqual({ scanVersion: '2', scanRunAsAdmin: false });
  });

  it('reads the program (exe) as a full version 2 scan', () => {
    expect(deriveScanInfo({ 'Scan Info': ['Version 2 exe | 2026-10-26 11:44 | Run by PC\\User | Admin No'] }))
      .toEqual({ scanVersion: '2', scanRunAsAdmin: false });
  });

  it('reads the basic fallback as "2 basic"', () => {
    expect(deriveScanInfo(fieldsOf(read('../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt'))))
      .toEqual({ scanVersion: '2 basic', scanRunAsAdmin: false });
  });

  it('reads a report with no Scan Info block as version 1', () => {
    expect(deriveScanInfo(fieldsOf(read('../__fixtures__/ASHRAF-PC_.txt'))))
      .toEqual({ scanVersion: '1', scanRunAsAdmin: null });
  });

  it('reads an administrator run', () => {
    const fields = { 'Scan Info': ['Version 2 | 2026-10-26 11:15 | Run by DOMAIN\\user | Admin Yes'] };
    expect(deriveScanInfo(fields)).toEqual({ scanVersion: '2', scanRunAsAdmin: true });
  });

  it('gives no version for a block it cannot read', () => {
    expect(deriveScanInfo({ 'Scan Info': ['Something else entirely'] }))
      .toEqual({ scanVersion: null, scanRunAsAdmin: null });
  });
});

describe('scanRank', () => {
  it('orders a full scan above a basic one, and both above version 1', () => {
    expect(scanRank('2')).toBeGreaterThan(scanRank('2 basic'));
    expect(scanRank('2 basic')).toBeGreaterThan(scanRank('1'));
    expect(scanRank('1')).toBeGreaterThan(scanRank(null));
  });

  it('ranks a row from before the field existed lowest', () => {
    expect(scanRank(null)).toBe(0);
    expect(scanRank(undefined)).toBe(0);
  });
});
