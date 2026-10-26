import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from './derive/deriveDevice.js';
import { licenceCards, licenceWarning, scanBadgeOf, scanKindOf, OLD_SCAN_SENTENCE } from './licenceView.js';
import { dayLabel, scanDayLabel } from './dayLabel.js';

const load = (path, fileName) => ({
  text: readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8'),
  fileName,
  lastModified: Date.UTC(2026, 9, 26, 3, 28),
});

const FULL = deriveDevice(load('./__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt', '[F1 IT] V2-SAMPLE_.txt'));
const BASIC = deriveDevice(load('./__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt', '[F3 FINANCE] V2-BASIC_.txt'));
const OLD = deriveDevice(load('./__fixtures__/ASHRAF-PC_.txt', 'ASHRAF-PC_.txt'));

const [WIN, OFF] = licenceCards(FULL);
const cardsWith = (patch) => licenceCards({ ...FULL, ...patch });

describe('dayLabel', () => {
  it('writes a stored date as day, short month, year', () => {
    expect(dayLabel('2027-01-21')).toBe('21 Jan 2027');
    expect(dayLabel('2026-10-05')).toBe('5 Oct 2026');
  });

  it('gives null for anything that is not a date', () => {
    expect(dayLabel(null)).toBeNull();
    expect(dayLabel('21/01/2027')).toBeNull();
    expect(dayLabel('2027-13-01')).toBeNull();
  });

  it('reads a scan instant as the Malaysia calendar day', () => {
    // 23:30 UTC on 26 Oct is already 27 Oct in Kuala Lumpur.
    expect(scanDayLabel(Date.UTC(2026, 9, 26, 23, 30))).toBe('27 Oct 2026');
    expect(scanDayLabel(undefined)).toBeNull();
  });
});

describe('Windows card, the states in Licences.dc.html', () => {
  it('OEM, activated: came with the PC and names the firmware edition', () => {
    const card = cardsWith({ windowsLicenseType: 'OEM', windowsActivation: 'Activated', firmwareKey: '[4.0] OEM:DM' })[0];
    expect(card).toMatchObject({ kicker: 'Windows', big: 'OEM', tone: 'Activated', toneKind: 'good' });
    expect(card.lines[0]).toBe('Came with the PC. Key in the firmware: Windows 11 Pro.');
  });

  it('Retail, activated: quotes the key ends', () => {
    const card = cardsWith({ windowsLicenseType: 'Retail', windowsActivation: 'Activated', windowsKeyEnds: '3V66T' })[0];
    expect(card).toMatchObject({ big: 'Retail', tone: 'Activated', toneKind: 'good' });
    expect(card.lines[0]).toBe('A bought key or a digital licence from an upgrade. Key ends 3V66T.');
  });

  it('Volume (MAK), activated: a company volume licence', () => {
    const card = cardsWith({ windowsLicenseType: 'Volume (MAK)', windowsActivation: 'Activated', windowsKeyEnds: '8XK2Q' })[0];
    expect(card).toMatchObject({ big: 'Volume (MAK)', tone: 'Activated', toneKind: 'good' });
    expect(card.lines[0]).toBe('Company volume licence. Key ends 8XK2Q.');
  });

  it('Volume (KMS), not activated: bad, and says why', () => {
    const card = cardsWith({ windowsLicenseType: 'Volume (KMS)', windowsActivation: 'Not activated' })[0];
    expect(card).toMatchObject({ big: 'Volume (KMS)', tone: 'Not activated', toneKind: 'bad' });
    expect(card.lines[0]).toBe('A KMS key with no KMS server reachable. Windows will nag and lock personalisation.');
  });

  it('Unknown licence on a full scan: says the scan did not read it', () => {
    const card = cardsWith({ windowsLicenseType: 'Unknown', windowsActivation: 'Unknown' })[0];
    expect(card).toMatchObject({ big: 'Unknown', tone: 'Not reported', toneKind: 'unknown' });
  });

  it('an old scan says so instead of guessing', () => {
    const card = licenceCards(OLD)[0];
    expect(card).toMatchObject({ big: 'Unknown', tone: 'Old scan', toneKind: 'unknown' });
    expect(card.lines).toEqual([OLD_SCAN_SENTENCE]);
  });

  it('the basic scan reads the licence like the full one', () => {
    expect(licenceCards(BASIC)[0]).toMatchObject({ big: 'Volume (KMS)', tone: 'Activated', toneKind: 'good' });
  });

  it('shows the edition and build as a fact under the state', () => {
    expect(WIN.lines).toContain('Windows 11 Pro · 26H2 (build 26300.9711, 64-bit)');
  });
});

describe('Office card, the states in Licences.dc.html', () => {
  it('Microsoft 365, company account only: good, with the check-in date', () => {
    const card = cardsWith({
      officeLicenseType: 'Microsoft 365', officePersonal: false, officePersonalReason: null,
      officeSignedInLicences: ['o365businessretail'], officeChannel: 'Current Channel',
      officeVersion: '16.0.20430.20182', officeLicensedTo: 'user@example.com',
      officeLicenceValidUntil: '2027-01-21',
    })[1];
    expect(card).toMatchObject({ big: 'Microsoft 365', tone: 'Company account', toneKind: 'good', warning: null });
    expect(card.lines).toContain('Business · Current Channel · 16.0.20430.20182');
    expect(card.lines).toContain('Checks in again by 21 Jan 2027');
  });

  it('Microsoft 365 with a personal copy beside the company one: warns, with the reason', () => {
    expect(OFF).toMatchObject({ big: 'Microsoft 365', tone: 'Personal copy beside it', toneKind: 'warn' });
    expect(OFF.warning).toBe(
      'A personal Microsoft 365 (O365HomePremRetail) is installed beside the company one. '
      + 'Remove the personal one so work files stay on the company account.',
    );
  });

  it('Microsoft 365 with only a personal account signed in: bad', () => {
    const card = cardsWith({
      officeLicenseType: 'Microsoft 365', officePersonal: true,
      officePersonalReason: 'A personal Microsoft account (me@outlook.com) is signed in to Office on this machine',
      officeSignedInLicences: ['o365homepremretail'],
    })[1];
    expect(card).toMatchObject({ tone: 'Personal account', toneKind: 'bad' });
    expect(card.warning).toMatch(/^A personal Microsoft account/);
  });

  it('Perpetual (volume), licensed: good', () => {
    const card = cardsWith({ officeLicenseType: 'Perpetual (volume)', officeVersion: '16.0.14332.20000' })[1];
    expect(card).toMatchObject({ big: 'Perpetual (volume)', tone: 'Licensed', toneKind: 'good' });
  });

  it('Perpetual (retail), licensed: good', () => {
    const card = cardsWith({ officeLicenseType: 'Perpetual (retail)' })[1];
    expect(card).toMatchObject({ big: 'Perpetual (retail)', tone: 'Licensed', toneKind: 'good' });
  });

  it('None: not installed, and free apps are not counted', () => {
    const card = cardsWith({ officeLicenseType: 'None', officePersonal: false })[1];
    expect(card).toMatchObject({ big: 'None', tone: 'Not installed', toneKind: 'unknown' });
    expect(card.lines[0]).toMatch(/Free apps \(Teams, OneNote\) are not counted/);
  });

  it('an old scan says so instead of guessing', () => {
    const card = licenceCards(OLD)[1];
    expect(card).toMatchObject({ big: 'Unknown', tone: 'Old scan', toneKind: 'unknown', warning: null });
    expect(card.lines).toEqual([OLD_SCAN_SENTENCE]);
  });

  it('a Microsoft 365 check-in date that was not stored leaves that line out', () => {
    const card = licenceCards(BASIC)[1];
    expect(card.lines.some((line) => line.startsWith('Checks in again'))).toBe(false);
  });
});

describe('Joined-to card, the states in Licences.dc.html', () => {
  it('Azure AD: good, and not on a domain', () => {
    expect(cardsWith({ deviceJoin: 'Azure AD' })[2]).toMatchObject({
      big: 'PMW Group', tone: 'Azure AD', toneKind: 'good',
    });
    expect(cardsWith({ deviceJoin: 'Azure AD' })[2].lines).toContain('Not on a domain');
  });

  it('Azure AD + domain: good', () => {
    expect(cardsWith({ deviceJoin: 'Azure AD + Domain' })[2]).toMatchObject({ tone: 'Azure AD + domain', toneKind: 'good' });
  });

  it('Domain only: good', () => {
    expect(cardsWith({ deviceJoin: 'Domain' })[2]).toMatchObject({ tone: 'Domain', toneKind: 'good' });
  });

  it('Registered only: warn', () => {
    expect(cardsWith({ deviceJoin: 'Registered only' })[2]).toMatchObject({ tone: 'Registered only', toneKind: 'warn' });
  });

  it('Not joined: bad, nothing to name', () => {
    expect(cardsWith({ deviceJoin: 'Not joined' })[2]).toMatchObject({ big: 'Nothing', tone: 'Not joined', toneKind: 'bad' });
  });

  it('Unknown, or an old scan: the Windows 7/8 wording', () => {
    expect(cardsWith({ deviceJoin: 'Unknown' })[2]).toMatchObject({ tone: 'Old scan or Windows 7/8', toneKind: 'unknown' });
    expect(licenceCards(OLD)[2]).toMatchObject({ tone: 'Old scan or Windows 7/8', toneKind: 'unknown' });
  });

  it('falls back to Company when the tenant name is not stored', () => {
    expect(cardsWith({ deviceJoin: 'Azure AD', tenantName: null })[2].big).toBe('Company');
  });
});

describe('licenceWarning', () => {
  it('is null when no card has a personal copy', () => {
    expect(licenceWarning({ ...FULL, officePersonal: false, officePersonalReason: null })).toBeNull();
  });

  it('is the personal-copy sentence on the full sample', () => {
    expect(licenceWarning(FULL)).toMatch(/^A personal Microsoft 365/);
  });

  it('is null on an old scan', () => {
    expect(licenceWarning(OLD)).toBeNull();
  });
});

describe('scanBadgeOf', () => {
  it('a full scan is Scan v2 with its day', () => {
    expect(scanBadgeOf(FULL)).toEqual({ kind: 'full', text: 'Scan v2 · 26 Oct 2026' });
    expect(scanKindOf(FULL)).toBe('full');
  });

  it('a basic scan says Basic scan', () => {
    expect(scanBadgeOf(BASIC)).toEqual({ kind: 'basic', text: 'Basic scan · 26 Oct 2026' });
  });

  it('an old scan, or a row with no version, says Old scan', () => {
    expect(scanBadgeOf(OLD)).toEqual({ kind: 'old', text: 'Old scan · 26 Oct 2026' });
    expect(scanBadgeOf({ ...FULL, scanVersion: null }).kind).toBe('old');
  });

  it('leaves the day off when no scan time was stored', () => {
    expect(scanBadgeOf({ ...FULL, scannedOn: null }).text).toBe('Scan v2');
  });
});
