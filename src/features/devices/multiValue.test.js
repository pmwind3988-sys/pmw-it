import { describe, it, expect } from 'vitest';
import { splitEntries, isMultiValue, entryCount } from './multiValue.js';

describe('splitEntries', () => {
  it('is empty for a blank field', () => {
    expect(splitEntries(null, 'gpuList')).toEqual([]);
    expect(splitEntries('', 'gpuList')).toEqual([]);
  });

  it('splits SharePoint newline-joined entries', () => {
    const entries = splitEntries('Intel UHD 620\nNVIDIA MX150', 'gpuList');
    expect(entries.map((e) => e.text)).toEqual(['Intel UHD 620', 'NVIDIA MX150']);
  });

  it('keeps a pipe as the parts of one entry, not as two entries', () => {
    const [entry] = splitEntries('Kingston | DDR4 | 3200 MHz', 'ramSlotInfoRaw');
    expect(entry.parts).toEqual(['Kingston', 'DDR4', '3200 MHz']);
  });

  it('splits Office on commas, because the report does', () => {
    const entries = splitEntries('Word, Excel, Outlook', 'microsoftOffice');
    expect(entries.map((e) => e.text)).toEqual(['Word', 'Excel', 'Outlook']);
  });

  it('leaves commas alone in a field that is one value', () => {
    const entries = splitEntries('Intel(R) Core(TM) i5-8250U, 4 cores', 'cpuModel');
    expect(entries).toHaveLength(1);
  });

  it('flattens the antivirus product objects', () => {
    const entries = splitEntries([{ product: 'Defender', enabled: true }], 'antivirusProducts');
    expect(entries[0].parts).toEqual(['Defender', 'Enabled']);
  });

  it('drops the blanks a trailing separator leaves behind', () => {
    expect(entryCount('Word,,Excel,', 'microsoftOffice')).toBe(2);
  });

  it('knows which fields hold several things', () => {
    expect(isMultiValue('gpuList')).toBe(true);
    expect(isMultiValue('cpuModel')).toBe(false);
  });
});
