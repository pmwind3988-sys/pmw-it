import { describe, it, expect } from 'vitest';
import { kindTagOf, backgroundCount, visibleSoftware, highlightParts, copyTextOf, installedLabel } from './softwareView.js';

const LIST = [
  { name: 'Google Chrome', version: '155.0.8059.85', publisher: 'Google LLC', installedOn: '2026-10-21' },
  { name: 'Microsoft Visual C++ 2015-2022 Redistributable (x64)', version: '14.40.33816', publisher: 'Microsoft Corporation', installedOn: null },
  { name: 'Microsoft Edge WebView2 Runtime', version: '130.0.2849', publisher: 'Microsoft Corporation', installedOn: null },
  { name: 'Realtek Audio Driver', version: '6.0.9549.1', publisher: 'Realtek Semiconductor Corp.', installedOn: null },
  { name: 'Microsoft Office 365 Apps for business', version: '16.0.20430', publisher: 'Microsoft Corporation', installedOn: null },
  { name: 'WPS Office', version: '12.2.0.13431', publisher: 'Kingsoft', installedOn: null },
];

describe('kindTagOf', () => {
  it('tags a runtime, a driver, and leaves a chosen program untagged', () => {
    expect(kindTagOf(LIST[1])).toBe('runtime');
    expect(kindTagOf(LIST[3])).toBe('driver');
    expect(kindTagOf(LIST[0])).toBeNull();
  });
});

describe('visibleSoftware', () => {
  it('folds runtimes and drivers away by default and counts them', () => {
    expect(backgroundCount(LIST)).toBe(3);
    expect(visibleSoftware(LIST).map((e) => e.name)).toEqual([
      'Google Chrome', 'Microsoft Office 365 Apps for business', 'WPS Office',
    ]);
  });

  it('shows everything when the switch is off', () => {
    expect(visibleSoftware(LIST, { hideRuntimes: false })).toHaveLength(6);
  });

  it('searches name and publisher, any case, and applies the switch too', () => {
    expect(visibleSoftware(LIST, { query: 'office' }).map((e) => e.name)).toEqual([
      'Microsoft Office 365 Apps for business', 'WPS Office',
    ]);
    expect(visibleSoftware(LIST, { query: 'kingsoft' })).toHaveLength(1);
    // A runtime matches the search but stays hidden until the switch is off.
    expect(visibleSoftware(LIST, { query: 'webview' })).toEqual([]);
    expect(visibleSoftware(LIST, { query: 'webview', hideRuntimes: false })).toHaveLength(1);
  });

  it('treats a blank search as no search', () => {
    expect(visibleSoftware(LIST, { query: '   ' })).toHaveLength(3);
  });
});

describe('highlightParts', () => {
  it('splits the name around every match, keeping the original case', () => {
    expect(highlightParts('Microsoft Office 365', 'office')).toEqual([
      { text: 'Microsoft ', match: false },
      { text: 'Office', match: true },
      { text: ' 365', match: false },
    ]);
  });

  it('returns the whole text unmatched when there is no search', () => {
    expect(highlightParts('Chrome', '')).toEqual([{ text: 'Chrome', match: false }]);
  });

  it('marks every occurrence', () => {
    const parts = highlightParts('aa-aa', 'aa');
    expect(parts.filter((p) => p.match)).toHaveLength(2);
  });
});

describe('copyTextOf and installedLabel', () => {
  it('copies one program per line as Name version', () => {
    expect(copyTextOf(LIST.slice(0, 2))).toBe(
      'Google Chrome 155.0.8059.85\nMicrosoft Visual C++ 2015-2022 Redistributable (x64) 14.40.33816',
    );
  });

  it('writes the install day, or a dash when none was recorded', () => {
    expect(installedLabel('2026-10-21')).toBe('21 Oct 2026');
    expect(installedLabel(null)).toBe('—');
  });
});
