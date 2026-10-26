import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from '../parse/parseReport.js';
import {
  deriveLicensing, windowsLicenseTypeOf, officeAccountsOf, parseOfficeAccounts, officeSubscriptionsOf,
} from './deriveLicensing.js';

const read = (path) => readFileSync(fileURLToPath(new URL(path, import.meta.url)), 'utf8');
const fieldsOf = (text) => parseReport(text).fields;

const FULL = fieldsOf(read('../__fixtures__/v2/[F1 IT] V2-SAMPLE_.txt'));
const BASIC = fieldsOf(read('../__fixtures__/v2/[F3 FINANCE] V2-BASIC_.txt'));
const OLD = fieldsOf(read('../__fixtures__/ASHRAF-PC_.txt'));

describe('windowsLicenseTypeOf', () => {
  it.each([
    ['OEM:DM', 'OEM'],
    ['OEM:NONSLP', 'OEM'],
    ['OEM_DM', 'OEM'],
    ['OEM_SLP', 'OEM'],
    ['Retail', 'Retail'],
    ['RETAIL', 'Retail'],
    ['Volume:MAK', 'Volume (MAK)'],
    ['VOLUME_MAK', 'Volume (MAK)'],
    ['Volume:GVLK', 'Volume (KMS)'],
    ['VOLUME_KMSCLIENT', 'Volume (KMS)'],
    ['Volume:CSVLK', 'Volume (KMS)'],
    ['Unknown', 'Unknown'],
    ['', 'Unknown'],
    [undefined, 'Unknown'],
  ])('reads the channel %j as %s', (channel, expected) => {
    expect(windowsLicenseTypeOf(channel)).toBe(expected);
  });
});

describe('deriveLicensing — full v2 scan', () => {
  const lic = deriveLicensing(FULL);

  it('reads Windows as a GVLK volume key, activated, with its key end and firmware edition', () => {
    expect(lic.windowsLicenseType).toBe('Volume (KMS)');
    expect(lic.windowsActivation).toBe('Activated');
    expect(lic.windowsKeyEnds).toBe('T83GX');
    expect(lic.firmwareKey).toBe('[4.0] CoreSingleLanguage OEM:DM');
  });

  it('reads Office as Microsoft 365, with the personal product flagged and why', () => {
    expect(lic.officeLicenseType).toBe('Microsoft 365');
    expect(lic.officePersonal).toBe(true);
    expect(lic.officePersonalReason).toBe(
      'A personal Microsoft 365 (O365HomePremRetail) is installed beside the company one',
    );
  });

  it('reads the signed-in licences, the account kind and the version', () => {
    expect(lic.officeSignedInLicences).toEqual(['o365businessretail', 'o365homepremretail']);
    expect(lic.officeAccounts).toEqual([{ who: null, kind: 'Work or school' }]);
    expect(lic.officeAccountsRaw).toBe('Work or school sign-in | 1 profile(s)');
    expect(lic.officeVersion).toBe('16.0.20430.20182');
    expect(lic.officeChannel).toBe('Current Channel');
  });

  it('reads the join state as Azure AD and the tenant', () => {
    expect(lic.deviceJoin).toBe('Azure AD');
    expect(lic.tenantName).toBe('PMW Group');
  });
});

describe('deriveLicensing — basic v2 scan', () => {
  const lic = deriveLicensing(BASIC);

  it('reads the KMS client key the cmd-only scan writes raw', () => {
    expect(lic.windowsLicenseType).toBe('Volume (KMS)');
    expect(lic.windowsActivation).toBe('Activated');
    expect(lic.firmwareKey).toBe(null); // "Firmware key | None"
  });

  it('still reads Office and the join state', () => {
    expect(lic.officeLicenseType).toBe('Microsoft 365');
    expect(lic.officePersonal).toBe(true);
    expect(lic.deviceJoin).toBe('Azure AD');
  });
});

describe('deriveLicensing — Office types', () => {
  const office = (licenseLines, accountLines = ['None']) =>
    deriveLicensing({ 'Office License': licenseLines, 'Office Accounts': accountLines }).officeLicenseType;

  it('says None for a scan that looked and found nothing', () => {
    expect(office(['None found'])).toBe('None');
  });

  it('says Perpetual (volume) for a volume key with no subscription', () => {
    expect(office(['ProPlus2021Volume | Perpetual, volume | Click-to-Run'])).toBe('Perpetual (volume)');
    expect(office(['Office 16, Office16ProPlus2021VL_KMS_Client edition | Volume | Licensed | Key ends ABCDE']))
      .toBe('Perpetual (volume)');
  });

  it('says Perpetual (retail) for a retail product with no subscription', () => {
    expect(office(['HomeBusiness2021Retail | Perpetual, retail | Click-to-Run'])).toBe('Perpetual (retail)');
  });

  it('lets Microsoft 365 win over a volume product beside it', () => {
    expect(office([
      'ProPlus2021Volume | Perpetual, volume | MSI',
      'O365BusinessRetail | Subscription (Microsoft 365) | Click-to-Run',
    ])).toBe('Microsoft 365');
  });

  it('says Unknown for an Office block it cannot classify', () => {
    expect(office(['Something | Odd | Thing'])).toBe('Unknown');
  });

  it('says Unknown for an old report with no Office block at all', () => {
    expect(deriveLicensing(OLD).officeLicenseType).toBe('Unknown');
  });
});

describe('deriveLicensing — personal Office', () => {
  it('is false for a company-only install with a work account', () => {
    const lic = deriveLicensing({
      'Office License': ['O365BusinessRetail | Subscription (Microsoft 365) | Click-to-Run'],
      'Office Accounts': ['user@pmw.example | Work or school'],
    });
    expect(lic.officePersonal).toBe(false);
    expect(lic.officePersonalReason).toBe(null);
  });

  it('is true for a personal Microsoft account signed in, and names it', () => {
    const lic = deriveLicensing({
      'Office License': ['O365BusinessRetail | Subscription (Microsoft 365) | Click-to-Run'],
      'Office Accounts': ['someone@outlook.example | Personal Microsoft account'],
    });
    expect(lic.officePersonal).toBe(true);
    expect(lic.officePersonalReason).toBe(
      'A personal Microsoft account (someone@outlook.example) is signed in to Office on this machine',
    );
  });

  it('is null for an old report, because nothing was looked at', () => {
    const lic = deriveLicensing(OLD);
    expect(lic.officePersonal).toBe(null);
    expect(lic.officePersonalReason).toBe(null);
  });
});

describe('deriveLicensing — Windows activation and channel', () => {
  it('says Not activated for any status other than Licensed', () => {
    const lic = deriveLicensing({
      'Windows License': ['Retail | Notification -- not activated | Key ends ABCDE | Windows(R), Home edition', 'Firmware key | None'],
    });
    expect(lic.windowsLicenseType).toBe('Retail');
    expect(lic.windowsActivation).toBe('Not activated');
    expect(lic.windowsKeyEnds).toBe('ABCDE');
  });

  it('says Unknown when the scan found no Windows key', () => {
    const lic = deriveLicensing({ 'Windows License': ['None found'] });
    expect(lic.windowsLicenseType).toBe('Unknown');
    expect(lic.windowsActivation).toBe('Unknown');
    expect(lic.windowsKeyEnds).toBe(null);
  });

  it('says Unknown for every Windows field of an old report', () => {
    const lic = deriveLicensing(OLD);
    expect(lic.windowsLicenseType).toBe('Unknown');
    expect(lic.windowsActivation).toBe('Unknown');
    expect(lic.firmwareKey).toBe(null);
  });
});

describe('deriveLicensing — device join', () => {
  const join = (lines) => deriveLicensing({ 'Device Join': lines });

  it('says Azure AD + Domain when both are yes', () => {
    expect(join(['Azure AD joined | YES', 'Domain joined | YES', 'Workplace joined | NO']).deviceJoin)
      .toBe('Azure AD + Domain');
  });

  it('says Domain when only the domain is yes', () => {
    expect(join(['Azure AD joined | NO', 'Domain joined | YES', 'Workplace joined | NO']).deviceJoin).toBe('Domain');
  });

  it('says Registered only when only the workplace join is yes', () => {
    expect(join(['Azure AD joined | NO', 'Domain joined | NO', 'Workplace joined | YES']).deviceJoin)
      .toBe('Registered only');
  });

  it('says Not joined when nothing is yes', () => {
    expect(join(['Azure AD joined | NO', 'Domain joined | NO', 'Workplace joined | NO']).deviceJoin)
      .toBe('Not joined');
  });

  it('says Unknown when the block is absent', () => {
    expect(deriveLicensing(OLD).deviceJoin).toBe('Unknown');
    expect(deriveLicensing(OLD).tenantName).toBe(null);
  });
});

describe('parseOfficeAccounts and officeAccountsOf', () => {
  it('reads an address with its kind, and a summary line with no address', () => {
    expect(parseOfficeAccounts(['a@b.example | Work or school', 'Personal Microsoft account sign-in | 2 profile(s)']))
      .toEqual([
        { who: 'a@b.example', kind: 'Work or school' },
        { who: null, kind: 'Personal Microsoft account' },
      ]);
  });

  it('parses a stored record that has only the raw lines', () => {
    const stored = { officeAccountsRaw: 'a@b.example | Work or school' };
    expect(officeAccountsOf(stored)).toEqual([{ who: 'a@b.example', kind: 'Work or school' }]);
  });

  it('returns the list of a fresh record untouched', () => {
    const fresh = deriveLicensing(FULL);
    expect(officeAccountsOf(fresh)).toBe(fresh.officeAccounts);
  });
});

describe('deriveLicensing — who Office is licensed to', () => {
  it('reads the account, tenant and next check-in from a full scan', () => {
    const d = deriveLicensing(FULL);
    expect(d.officeLicensedTo).toBe('user@example.com');
    expect(d.officeTenantId).toBe('3042ec28-18ef-448b-a1fd-cf1f2446943c');
    expect(d.officeLicenceValidUntil).toBe('2027-01-21');
    expect(d.officeSubscriptions).toEqual([{
      product: 'O365BusinessRetail', type: 'Subscription', account: 'user@example.com',
      tenant: '3042ec28-18ef-448b-a1fd-cf1f2446943c', validUntil: '2027-01-21', status: 'Provisioned', trial: false,
    }]);
  });

  it('keeps only the account from a basic scan, never the word Unknown', () => {
    const d = deriveLicensing(BASIC);
    expect(d.officeLicensedTo).toBe('user@example.com');
    expect(d.officeTenantId).toBeNull();
    expect(d.officeLicenceValidUntil).toBeNull();
  });

  it('says nothing for an old report', () => {
    const d = deriveLicensing(OLD);
    expect(d.officeLicensedTo).toBeNull();
    expect(d.officeSubscriptions).toEqual([]);
    expect(d.officeSubscriptionsRaw).toBeNull();
  });

  it('shows the company licence when a personal one is on the PC too', () => {
    const d = deriveLicensing({ 'Office Subscription': [
      'O365HomePremRetail | Subscription | me@outlook.com | Tenant 9188040d-6c67-4c5b-b112-36a304b66dad | Valid until 2026-12-01 | Provisioned',
      'O365BusinessRetail | Subscription | ali@pmw-group.com | Tenant 3042ec28 | Valid until 2027-01-21 | Provisioned | Trial',
    ] });
    expect(d.officeLicensedTo).toBe('ali@pmw-group.com');
    expect(d.officeSubscriptions[1].trial).toBe(true);
  });

  it('reads the stored raw lines back the same way', () => {
    const fresh = deriveLicensing(FULL);
    expect(officeSubscriptionsOf({ officeSubscriptionsRaw: fresh.officeSubscriptionsRaw })).toEqual(fresh.officeSubscriptions);
  });
});
