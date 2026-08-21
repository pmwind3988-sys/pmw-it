import { describe, it, expect } from 'vitest';
import { deriveHealth } from './deriveHealth.js';

const base = {
  'Computer Name': ['PC1'], Processor: ['Intel'], 'Storage Drives': ['D | SSD | 477 GB'],
  'Windows Version': [], 'Antivirus status': [], Antivirus: [],
};
const withFields = (overrides) => ({ ...base, ...overrides });

describe('deriveHealth — Windows', () => {
  it('reads major version and edition', () => {
    const result = deriveHealth(withFields({ 'Windows Version': ['Microsoft Windows 11 Pro'] }));
    expect(result.windowsMajor).toBe(11);
    expect(result.windowsEdition).toBe('Pro');
    expect(result.osSupported).toBe(true);
  });

  it('reads the long Home edition name', () => {
    const result = deriveHealth(withFields({
      'Windows Version': ['Microsoft Windows 11 Home Single Language'],
    }));
    expect(result.windowsEdition).toBe('Home Single Language');
  });

  it('marks Windows 10 unsupported — it lost support on 14 October 2025', () => {
    const result = deriveHealth(withFields({ 'Windows Version': ['Microsoft Windows 10 Pro'] }));
    expect(result.windowsMajor).toBe(10);
    expect(result.osSupported).toBe(false);
  });

  it('returns null support for a report with no Windows line', () => {
    expect(deriveHealth(withFields({})).osSupported).toBe(null);
  });
});

describe('deriveHealth — antivirus status', () => {
  const status = (raw, products = []) =>
    deriveHealth(withFields({ 'Antivirus status': raw ? [raw] : [], Antivirus: products }))
      .antivirusStatus;

  it('normalises every spelling the scan produces', () => {
    expect(status('NORTON NOT INSTALLED')).toBe('Not Installed');
    expect(status('NORTON ACTIVATED')).toBe('Active');
    expect(status('NORTON ACTIVE')).toBe('Active');
    expect(status('NORTON INSTALLED (ACTIVE)')).toBe('Active');
    expect(status('NORTON INSTALLED (7 DAYS)')).toBe('Trial');
  });

  it('does not read DEACTIVATED as active', () => {
    expect(status('NORTON INSTALLED (DEACTIVATED)')).toBe('Installed — Inactive');
  });

  it('falls back to the antivirus block when the status line is blank', () => {
    expect(status('', ['Norton 360 | Enabled'])).toBe('Active');
    expect(status('', ['Norton 360 | Disabled'])).toBe('Installed — Inactive');
    expect(status('', [])).toBe('Unknown');
  });
});

describe('deriveHealth — protection', () => {
  it('counts Windows Defender as protection', () => {
    const result = deriveHealth(withFields({
      'Antivirus status': ['NORTON NOT INSTALLED'],
      Antivirus: ['Windows Defender | Enabled'],
    }));
    expect(result.avProtected).toBe(true);
  });

  it('reports unprotected when every product is disabled', () => {
    const result = deriveHealth(withFields({
      Antivirus: ['Norton 360 | Disabled', 'Windows Defender | Disabled'],
    }));
    expect(result.avProtected).toBe(false);
  });

  it('de-duplicates the repeated products before judging', () => {
    const result = deriveHealth(withFields({
      Antivirus: Array(22).fill('HP Wolf Pro Security | Enabled'),
    }));
    expect(result.antivirusProducts).toEqual([{ product: 'HP Wolf Pro Security', enabled: true }]);
  });
});

describe('deriveHealth — scan completeness', () => {
  it('is complete when the core fields are present', () => {
    expect(deriveHealth(base).scanComplete).toBe(true);
  });

  it('is incomplete when name, processor and storage are all empty', () => {
    const result = deriveHealth({
      ...base, 'Computer Name': [], Processor: [], 'Storage Drives': [],
    });
    expect(result.scanComplete).toBe(false);
  });

  it('is complete when only one core field is missing', () => {
    expect(deriveHealth(withFields({ Processor: [] })).scanComplete).toBe(true);
  });
});
