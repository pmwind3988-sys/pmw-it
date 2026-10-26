import { cleanValue } from '../parse/placeholders.js';
import { splitParts, cellValue, linesOf, meaningful } from '../parse/parts.js';

/**
 * What the machine's licences say, read from the scan version 2 blocks:
 * Windows (how it was bought, whether it is activated), Office (which products
 * and how they are licensed, and whose account is signed in), and whether the
 * machine belongs to the company's Microsoft tenant.
 *
 * These describe what is installed and signed in, not what the company paid
 * for. The Authentic / Unlicensed judgement that existed before stays in
 * officeLicense.js; nothing here replaces it.
 *
 * An old report has none of these blocks. Every value then reads 'Unknown' or
 * null, except the list fields, which read empty.
 */

export const WINDOWS_LICENSE_TYPES = ['OEM', 'Retail', 'Volume (MAK)', 'Volume (KMS)', 'Unknown'];
export const OFFICE_LICENSE_TYPES = ['Microsoft 365', 'Perpetual (volume)', 'Perpetual (retail)', 'None', 'Unknown'];
export const DEVICE_JOIN_STATES = ['Azure AD', 'Domain', 'Azure AD + Domain', 'Registered only', 'Not joined', 'Unknown'];

/** Consumer Office: a home or personal subscription or one-off licence, never the company's. */
const CONSUMER = /HomePrem|HomeBusiness|HomeStudent|Personal/i;
const MICROSOFT_365 = /O365|M365/i;

/**
 * The Windows channel, in either spelling the scan has produced: `OEM:DM`,
 * `OEM_SLP`, `Volume:GVLK`, `VOLUME_KMSCLIENT`, `Retail`, `Volume:MAK`...
 * Anything else is Unknown rather than a guess.
 */
export function windowsLicenseTypeOf(channel) {
  const text = String(channel ?? '').toUpperCase().replace(/[:_]/g, ' ').trim();
  if (text.startsWith('OEM')) return 'OEM';
  if (text.startsWith('RETAIL')) return 'Retail';
  if (text.startsWith('VOLUME')) {
    // MAK is checked first: a MAK key is a company key that is activated once and
    // stays activated. GVLK and CSVLK are KMS keys, which need a KMS host on the network.
    if (text.includes('MAK')) return 'Volume (MAK)';
    if (/GVLK|KMS|CSVLK/.test(text)) return 'Volume (KMS)';
  }
  return 'Unknown';
}

export function deriveWindowsLicense(fields) {
  const lines = fields['Windows License'] ?? [];
  const firmware = lines.map(splitParts).find((cells) => /^Firmware key$/i.test(cells[0]));
  const licence = meaningful(lines).find((line) => !/^Firmware key/i.test(line));
  const cells = licence ? splitParts(licence) : null;

  let windowsActivation = 'Unknown';
  if (cells && cells[1]) windowsActivation = /^licensed$/i.test(cells[1]) ? 'Activated' : 'Not activated';

  return {
    windowsLicenseType: cells ? windowsLicenseTypeOf(cells[0]) : 'Unknown',
    windowsActivation,
    windowsKeyEnds: cells ? cellValue(cells, /^Key ends\s*/i) : null,
    // The edition the maker burned into the firmware. Only the edition is written, never the key.
    firmwareKey: firmware ? cleanValue(firmware[1]) : null,
  };
}

/** `user@example.com | Work or school`, or the summary `Work or school sign-in | 1 profile(s)`. */
export function parseOfficeAccounts(lines) {
  return meaningful(lines).map((line) => {
    const cells = splitParts(line);
    if (cells[0].includes('@')) {
      return { who: cleanValue(cells[0]), kind: cleanValue(cells[1]) };
    }
    // A summary line names the kind of sign-in only; the address is not kept.
    return { who: null, kind: cleanValue(cells[0].replace(/\s+sign-in$/i, '')) };
  });
}

function parseOfficeLicenses(lines) {
  const products = [];
  const licenceLines = [];
  const signedIn = [];

  for (const line of meaningful(lines)) {
    const cells = splitParts(line);
    if (/^Signed-in licence$/i.test(cells[0])) {
      const id = cleanValue(cells[1]);
      if (id) signedIn.push(id);
    } else if (/^Office\b/i.test(cells[0])) {
      // `Office 16, Office16O365BusinessR_Grace edition | Retail | Notification -- not activated | Key ends 3RQ6B`
      licenceLines.push({ name: cells[0], channel: cells[1] ?? '' });
    } else if (cleanValue(cells[0])) {
      // `O365BusinessRetail | Subscription (Microsoft 365) | Click-to-Run`
      products.push({ id: cleanValue(cells[0]) });
    }
  }

  // Products first, then the signed-in ones that are not already named by a product.
  const ids = [...products.map((p) => p.id)];
  for (const id of signedIn) {
    if (!ids.some((known) => known.toLowerCase() === id.toLowerCase())) ids.push(id);
  }

  return { products, licenceLines, signedIn, ids };
}

export function deriveOffice(fields) {
  const licenseLines = fields['Office License'] ?? [];
  const accountLines = fields['Office Accounts'] ?? [];
  const versionLines = meaningful(fields['Office Version']);
  const accountsRaw = meaningful(accountLines);

  // A block that is absent (not even `None found`) means an old report.
  const hasLicenseBlock = licenseLines.length > 0;
  const hasAccountBlock = accountLines.length > 0;

  const { products, licenceLines, signedIn, ids } = parseOfficeLicenses(licenseLines);
  const accounts = parseOfficeAccounts(accountLines);

  let officeLicenseType = 'Unknown';
  if (hasLicenseBlock) {
    const anything = products.length || licenceLines.length || signedIn.length;
    if (!anything) officeLicenseType = 'None';
    else if (ids.some((id) => MICROSOFT_365.test(id))
      || licenceLines.some((line) => MICROSOFT_365.test(line.name))) officeLicenseType = 'Microsoft 365';
    else if (products.some((p) => /Volume$/i.test(p.id))
      || licenceLines.some((line) => /volume/i.test(line.channel))) officeLicenseType = 'Perpetual (volume)';
    else if (products.some((p) => /Retail$/i.test(p.id))
      || licenceLines.some((line) => /retail/i.test(line.channel))) officeLicenseType = 'Perpetual (retail)';
  }

  // Null for an old report: nothing was looked at, so nothing can be said either way.
  let officePersonal = null;
  let officePersonalReason = null;
  if (hasLicenseBlock || hasAccountBlock) {
    const consumerIds = ids.filter((id) => CONSUMER.test(id));
    const personalAccount = accounts.find((account) => /personal/i.test(account.kind ?? ''));
    officePersonal = consumerIds.length > 0 || Boolean(personalAccount);

    const reasons = [];
    const companyPresent = ids.some((id) => !CONSUMER.test(id));
    for (const id of consumerIds) {
      const label = MICROSOFT_365.test(id) ? 'Microsoft 365' : 'Office';
      reasons.push(`A personal ${label} (${id}) is installed${companyPresent ? ' beside the company one' : ''}`);
    }
    if (personalAccount) {
      const who = personalAccount.who ? ` (${personalAccount.who})` : '';
      reasons.push(`A personal Microsoft account${who} is signed in to Office on this machine`);
    }
    officePersonalReason = reasons.length ? reasons.join('; ') : null;
  }

  const [version] = versionLines.map(splitParts);

  return {
    officeLicenseType,
    officePersonal,
    officePersonalReason,
    officeSignedInLicences: signedIn,
    officeAccounts: accounts,
    officeAccountsRaw: accountsRaw.length ? accountsRaw.join('\n') : null,
    officeVersion: cleanValue(version?.[0]),
    officeChannel: cleanValue(version?.[2]),
  };
}

export function deriveDeviceJoin(fields) {
  const lines = fields['Device Join'] ?? [];
  const cells = lines.map(splitParts);
  const value = (label) => {
    const hit = cells.find((entry) => entry[0].toLowerCase() === label);
    return hit ? hit[1] : undefined;
  };

  const aad = value('azure ad joined');
  const domain = value('domain joined');
  const workplace = value('workplace joined');
  const tenant = cleanValue(value('tenant'));

  if (aad === undefined && domain === undefined && workplace === undefined) {
    return { deviceJoin: 'Unknown', tenantName: tenant };
  }

  const isYes = (state) => /^yes$/i.test(state ?? '');
  let deviceJoin = 'Not joined';
  if (isYes(aad) && isYes(domain)) deviceJoin = 'Azure AD + Domain';
  else if (isYes(aad)) deviceJoin = 'Azure AD';
  else if (isYes(domain)) deviceJoin = 'Domain';
  else if (isYes(workplace)) deviceJoin = 'Registered only';

  return { deviceJoin, tenantName: tenant };
}

/**
 * `O365BusinessRetail | Subscription | user@pmw-group.com | Tenant <id> | Valid until 2027-01-21 | Provisioned`
 * One line per Microsoft 365 licence on the PC. "Valid until" is the next time
 * the PC must check the licence in online, NOT the end of the paid
 * subscription; the plan itself (Business Basic, Standard, Premium) is not on
 * the PC at all. The basic scan only knows the account, so the rest is Unknown.
 */
export function parseOfficeSubscriptions(lines) {
  const unknown = (value) => {
    const v = cleanValue(value);
    return v && !/^unknown( account)?$/i.test(v) ? v : null;
  };
  return meaningful(lines).map((line) => {
    const cells = splitParts(line);
    return {
      product: unknown(cells[0]),
      type: unknown(cells[1]),
      account: unknown(cells[2]),
      tenant: unknown(cellValue(cells, /^Tenant\s*/i)),
      validUntil: unknown(cellValue(cells, /^Valid until\s*/i)),
      status: unknown(cells[5]),
      trial: cells.some((cell) => /^trial$/i.test(cell)),
    };
  });
}

export function deriveOfficeSubscription(fields) {
  const lines = meaningful(fields['Office Subscription']);
  const subscriptions = parseOfficeSubscriptions(lines);
  // The company licence is the one to show: a work account first, then any.
  const main = subscriptions.find((s) => s.account && !CONSUMER.test(s.product ?? ''))
    ?? subscriptions.find((s) => s.account) ?? subscriptions[0];
  return {
    officeLicensedTo: main?.account ?? null,
    officeTenantId: main?.tenant ?? null,
    officeLicenceValidUntil: main?.validUntil ?? null,
    officeSubscriptions: subscriptions,
    officeSubscriptionsRaw: lines.length ? lines.join('\n') : null,
  };
}

export function deriveLicensing(fields) {
  return {
    ...deriveWindowsLicense(fields),
    ...deriveOffice(fields),
    ...deriveOfficeSubscription(fields),
    ...deriveDeviceJoin(fields),
  };
}

/** Every licence line as objects, for a stored record (raw text only) or a fresh one. */
export function officeSubscriptionsOf(device) {
  if (Array.isArray(device?.officeSubscriptions)) return device.officeSubscriptions;
  return parseOfficeSubscriptions(linesOf(device?.officeSubscriptionsRaw));
}

/** The signed-in account list as objects, for a stored record (raw text only) or a fresh one. */
export function officeAccountsOf(device) {
  if (Array.isArray(device?.officeAccounts)) return device.officeAccounts;
  return parseOfficeAccounts(linesOf(device?.officeAccountsRaw));
}
