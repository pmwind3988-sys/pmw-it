import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from './parseReport.js';

const fixture = (name) =>
  readFileSync(fileURLToPath(new URL(`../__fixtures__/${name}`, import.meta.url)), 'utf8');

describe('parseReport — structure', () => {
  const parsed = parseReport(fixture('ASHRAF-PC_.txt'));

  it('reads a single-answer field', () => {
    expect(parsed.fields['Computer Name']).toEqual(['ASHRAF-PC']);
  });

  it('reads a multi-answer field as every line', () => {
    expect(parsed.fields['GPU']).toEqual([
      'Intel(R) Iris(R) Xe Graphics',
      'VirtualMonitorDriver Device',
    ]);
  });

  it('keeps the RAM slot summary inside the RAM Slot Info block', () => {
    expect(parsed.fields['RAM Slot Info']).toEqual([
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      'Total Slots: 2 | Used Slots: 2',
    ]);
  });

  it('keeps mapped drive lines inside Server folder', () => {
    expect(parsed.fields['Server folder']).toEqual([
      'Y: | \\\\server\\emdata$\\device list 2026',
      'Z: | \\\\server\\PMW\\IT',
    ]);
  });

  it('does not leak the banner text into Remarks', () => {
    expect(parsed.fields['Remarks']).toEqual([]);
  });

  it('records no unknown labels for a standard report', () => {
    expect(parsed.unknownLabels).toEqual([]);
  });
});

describe('parseReport — inline values', () => {
  it('reads an inline value written with a space after the colon', () => {
    const parsed = parseReport(fixture('[ENGINEERING] AMIR-HP_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON ACTIVE']);
  });

  it('reads an inline value written with no space after the colon', () => {
    const parsed = parseReport(fixture('[QAQC FAIRUS]HPFL05_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON INSTALLED (ACTIVE)']);
  });
});

describe('parseReport — CRLF and blank reports', () => {
  it('handles CRLF line endings without leaving stray returns', () => {
    const parsed = parseReport(fixture('CARMEN-HP_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON INSTALLED (7 DAYS)']);
  });

  it('parses a report whose every field is empty', () => {
    const parsed = parseReport(fixture('CARMEN-HP_.txt'));
    expect(parsed.isReport).toBe(true);
    expect(parsed.fields['Computer Name']).toEqual([]);
    expect(parsed.fields['Processor']).toEqual([]);
  });
});

describe('parseReport — not a report', () => {
  it('flags a file with no known label', () => {
    const parsed = parseReport('Dear team,\n\nPlease find the invoice attached.\n');
    expect(parsed.isReport).toBe(false);
  });
});

describe('parseReport — unknown labels', () => {
  it('records a label the scan script did not used to write', () => {
    const parsed = parseReport('Computer Name:\nPC1\n\nBitLocker Status:\nEnabled\n');
    expect(parsed.unknownLabels).toEqual([{ label: 'BitLocker Status', value: 'Enabled' }]);
  });

  it('does not record a pipe-delimited value as an unknown label', () => {
    const parsed = parseReport('RAM Slot Info:\nTotal Slots: 2 | Used Slots: 2\n');
    expect(parsed.unknownLabels).toEqual([]);
  });

  it('does not record a drive letter as an unknown label', () => {
    const parsed = parseReport('Server folder:\nY: | \\\\server\\PMW\n');
    expect(parsed.unknownLabels).toEqual([]);
  });

  it('does not record a Windows path as an unknown label', () => {
    const parsed = parseReport('Remarks:\nC:\\Users\\User\\Desktop\n');
    expect(parsed.unknownLabels).toEqual([]);
  });
});

describe('parseReport — blank lines inside a block', () => {
  it('does not truncate a block at a blank line', () => {
    const parsed = parseReport('GPU:\nIntel HD\n\nNVIDIA RTX\n\nTotal RAM:\n8 GB\n');
    expect(parsed.fields['GPU']).toEqual(['Intel HD', 'NVIDIA RTX']);
    expect(parsed.fields['Total RAM']).toEqual(['8 GB']);
  });
});
