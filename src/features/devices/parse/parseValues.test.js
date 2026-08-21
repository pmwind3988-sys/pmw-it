import { describe, it, expect } from 'vitest';
import {
  parseSize, parseRamSlots, parseDrives, parseNetwork, parseAntivirus,
  parsePairs, parseOffice, parseGpus, parseMonitors, parseMailFiles,
} from './parseValues.js';

describe('parseSize', () => {
  it('reads GB and TB', () => {
    expect(parseSize('477 GB')).toBe(477);
    expect(parseSize('8 GB')).toBe(8);
    expect(parseSize('1 TB')).toBe(1024);
  });
  it('returns null for anything unparseable', () => {
    expect(parseSize('')).toBe(null);
    expect(parseSize('Unknown')).toBe(null);
  });
});

describe('parseRamSlots', () => {
  it('reads two sticks and the summary line', () => {
    const result = parseRamSlots([
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      'Total Slots: 2 | Used Slots: 2',
    ]);
    expect(result.sticks).toHaveLength(2);
    expect(result.sticks[0]).toEqual({
      sizeGB: 4, type: 'DDR4', speedMhz: 3200,
      vendor: 'Samsung', partNumber: 'M471A5244CB0-CWE',
    });
    expect(result.totalSlots).toBe(2);
    expect(result.usedSlots).toBe(2);
  });

  it('falls back to counting sticks when Used Slots is blank', () => {
    const result = parseRamSlots([
      '16 GB | DDR4 | 3200 MHz | Samsung | M471A2G43AB2-CWE',
      'Total Slots: 2 | Used Slots: ',
    ]);
    expect(result.totalSlots).toBe(2);
    expect(result.usedSlots).toBe(1);
  });

  it('nulls the unset SMBIOS vendor and part number', () => {
    const result = parseRamSlots([
      '2 GB | Unknown | 333 MHz | Manufacturer1 | PartNum1',
      'Total Slots: 2 | Used Slots: ',
    ]);
    expect(result.sticks[0].vendor).toBe(null);
    expect(result.sticks[0].partNumber).toBe(null);
    expect(result.sticks[0].type).toBe(null);
    expect(result.sticks[0].sizeGB).toBe(2);
  });

  it('never counts the summary line as a stick', () => {
    const result = parseRamSlots(['Total Slots: 4 | Used Slots: ']);
    expect(result.sticks).toEqual([]);
    expect(result.usedSlots).toBe(0);
  });
});

describe('parseDrives', () => {
  it('reads model, type and size', () => {
    expect(parseDrives(['KBG50ZNV512G KIOXIA | SSD | 477 GB'])).toEqual([
      { model: 'KBG50ZNV512G KIOXIA', type: 'SSD', sizeGB: 477, mechanical: false },
    ]);
  });

  it('treats Unspecified as a mechanical disk', () => {
    const [drive] = parseDrives(['WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB']);
    expect(drive.type).toBe('HDD (assumed)');
    expect(drive.mechanical).toBe(true);
  });
});

describe('parseNetwork', () => {
  it('strips the inner SSID and IP prefixes', () => {
    expect(parseNetwork(['Wi-Fi | SSID: PMW_Group 7 | IP: 192.168.1.170 | Dynamic'])).toEqual({
      connection: 'Wi-Fi', ssid: 'PMW_Group 7', ip: '192.168.1.170', assignment: 'Dynamic',
    });
  });
  it('returns null when the block is empty', () => {
    expect(parseNetwork([])).toBe(null);
  });
});

describe('parseAntivirus', () => {
  it('de-duplicates repeated products and keeps enabled if any entry is enabled', () => {
    const result = parseAntivirus([
      'HP Wolf Pro Security | Enabled',
      'HP Wolf Pro Security | Disabled',
      'HP Wolf Pro Security | Enabled',
      'Norton 360 | Enabled',
      'Windows Defender | Disabled',
    ]);
    expect(result).toEqual([
      { product: 'HP Wolf Pro Security', enabled: true },
      { product: 'Norton 360', enabled: true },
      { product: 'Windows Defender', enabled: false },
    ]);
  });
});

describe('parsePairs', () => {
  it('splits a two-part line', () => {
    expect(parsePairs(['HP | 8BB6'])).toEqual([{ left: 'HP', right: '8BB6' }]);
  });
  it('keeps the whole line as left when there is no pipe', () => {
    expect(parsePairs(['server'])).toEqual([{ left: 'server', right: null }]);
  });
});

describe('parseOffice', () => {
  it('splits the single comma-separated line', () => {
    expect(parseOffice(['O365BusinessRetail,O365HomePremRetail']))
      .toEqual(['O365BusinessRetail', 'O365HomePremRetail']);
  });
});

describe('parseGpus', () => {
  it('drops the AnyDesk virtual display', () => {
    expect(parseGpus(['Intel(R) Iris(R) Xe Graphics', 'VirtualMonitorDriver Device']))
      .toEqual(['Intel(R) Iris(R) Xe Graphics']);
  });
});

describe('parseMonitors', () => {
  it('drops the Windows pseudo-monitor', () => {
    expect(parseMonitors(['Generic PnP Monitor', 'Default Monitor']))
      .toEqual(['Generic PnP Monitor']);
  });
});

describe('parseMailFiles', () => {
  it('classifies .ost as a mailbox and .pst as an archive', () => {
    const result = parseMailFiles([
      'ashraf@pmw-group.com.ost | C:\\Users\\User\\AppData\\Local\\Microsoft\\Outlook\\a.ost',
      'ashraf@pmw-industries.com.pst | C:\\Users\\User\\Documents\\Outlook Files\\b.pst',
    ]);
    expect(result.map((r) => r.kind)).toEqual(['mailbox', 'archive']);
  });
});
