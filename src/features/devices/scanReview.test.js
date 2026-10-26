import { describe, expect, it } from 'vitest';
import { isOlderThanStored, olderScanNote, scanSummary, scanSummaryLine } from './scanReview.js';

const day = (text) => Date.parse(`${text}T02:00:00+08:00`);

const matches = [
  { device: { scanVersion: '2' }, existing: null },
  { device: { scanVersion: '2 basic' }, existing: { scanVersion: '2', scannedOn: day('2026-10-03') } },
  { device: { scanVersion: '1' }, existing: { scanVersion: '2', scannedOn: day('2026-10-03') } },
  { device: { scanVersion: null }, existing: null },
  { device: { scanVersion: '2' }, existing: { scanVersion: '2 basic', scannedOn: day('2026-09-30') } },
];

describe('scanSummary', () => {
  it('counts each kind of report and the ones older than what is stored', () => {
    expect(scanSummary(matches)).toEqual({
      reports: 5, new: 2, basic: 1, old: 2, olderThanStored: 2,
    });
  });

  it('counts nothing as older when there is nothing stored', () => {
    expect(scanSummary([{ device: { scanVersion: '1' }, existing: null }]).olderThanStored).toBe(0);
  });
});

describe('isOlderThanStored', () => {
  it('follows the rank the import keeps licences by', () => {
    expect(isOlderThanStored('2 basic', '2')).toBe(true);
    expect(isOlderThanStored('2', '2 basic')).toBe(false);
    expect(isOlderThanStored('2', '2')).toBe(false);
    expect(isOlderThanStored('1', null)).toBe(false);
  });
});

describe('scanSummaryLine', () => {
  it('reads the summary in plain words, singular for one report', () => {
    expect(scanSummaryLine(scanSummary(matches)))
      .toBe('5 reports read · 2 new scan · 1 basic scan · 2 old scan · 2 older than what is stored — licences and software kept');
    expect(scanSummaryLine(scanSummary([{ device: { scanVersion: '2' }, existing: null }])))
      .toMatch(/^1 report read · 1 new scan/);
  });
});

describe('olderScanNote', () => {
  it('names the stored scan, its version and day, when the incoming one is older', () => {
    expect(olderScanNote(matches[1])).toBe('Stored scan is newer (v2, 3 Oct). Licences and software kept.');
  });

  it('says nothing when the incoming scan is not older', () => {
    expect(olderScanNote(matches[0])).toBeNull();
    expect(olderScanNote(matches[4])).toBeNull();
  });
});
