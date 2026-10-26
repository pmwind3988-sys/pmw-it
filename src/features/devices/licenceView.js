import { dayLabel, scanDayLabel } from './dayLabel.js';

/**
 * What the three licence cards on a machine page say, and the badge that says
 * which scan the page is built from. Pure: the page draws whatever this returns.
 *
 * Each card is { kicker, big, tone, toneKind, lines, warning }:
 *   tone      the words on the pill; the words carry the meaning, colour only adds to them
 *   toneKind  'good' | 'warn' | 'bad' | 'unknown', which grade colour the pill takes
 *   lines     the explanation first, then the facts under it
 *   warning   the sentence for the amber box under the cards, or null
 *
 * An old scan (version 1, or a row from before the field existed) has none of
 * the licence blocks, so it never guesses: it says so.
 */

export const OLD_SCAN_SENTENCE = 'This machine was last scanned with the old scan. Scan it again to see its licence.';

/** Consumer Office: the same pattern deriveLicensing uses to call a copy personal. */
const CONSUMER = /HomePrem|HomeBusiness|HomeStudent|Personal/i;

export function isOldScan(device) {
  const version = device?.scanVersion;
  return version === '1' || version === null || version === undefined || version === '';
}

/** Which scan the page is built from: 'full', 'basic' or 'old'. */
export function scanKindOf(device) {
  if (isOldScan(device)) return 'old';
  return /basic/i.test(String(device.scanVersion)) ? 'basic' : 'full';
}

/** The header badge: "Scan v2 · 26 Oct 2026", "Basic scan · ..." or "Old scan · ...". */
export function scanBadgeOf(device) {
  const kind = scanKindOf(device);
  const day = scanDayLabel(device?.scannedOn);
  const label = kind === 'full' ? 'Scan v2' : kind === 'basic' ? 'Basic scan' : 'Old scan';
  return { kind, text: day ? `${label} · ${day}` : label };
}

const windowsEdition = (version) => {
  const text = String(version ?? '').replace(/^Microsoft\s+/i, '').trim();
  return text || null;
};

const keyLine = (key) => (key ? `Key ends ${key}.` : '');

const oldCard = (kicker) => ({
  kicker, big: 'Unknown', tone: 'Old scan', toneKind: 'unknown', lines: [OLD_SCAN_SENTENCE], warning: null,
});

function windowsCard(device) {
  const kicker = 'Windows';
  if (isOldScan(device)) return oldCard(kicker);

  const type = device.windowsLicenseType ?? 'Unknown';
  const activation = device.windowsActivation ?? 'Unknown';
  const edition = windowsEdition(device.windowsVersion);
  const facts = [[edition, device.windowsBuild].filter(Boolean).join(' · ')].filter(Boolean);

  if (type === 'Unknown') {
    return {
      kicker, big: 'Unknown', tone: 'Not reported', toneKind: 'unknown',
      lines: ['This scan did not read the Windows licence.', ...facts], warning: null,
    };
  }

  const notActivated = activation === 'Not activated';
  const tone = notActivated ? 'Not activated' : activation === 'Activated' ? 'Activated' : 'Activation not reported';
  const toneKind = notActivated ? 'bad' : activation === 'Activated' ? 'good' : 'unknown';

  let say;
  if (type === 'OEM') {
    say = device.firmwareKey
      ? `Came with the PC. Key in the firmware: ${edition ?? 'Windows'}.`
      : 'Came with the PC. No key in the firmware.';
  } else if (type === 'Retail') {
    say = `A bought key or a digital licence from an upgrade. ${keyLine(device.windowsKeyEnds)}`.trim();
  } else if (type === 'Volume (MAK)') {
    say = `Company volume licence. ${keyLine(device.windowsKeyEnds)}`.trim();
  } else if (notActivated) {
    // A KMS key with no server to check it against: the nag and the locked personalisation.
    say = 'A KMS key with no KMS server reachable. Windows will nag and lock personalisation.';
  } else {
    say = `Company KMS key. ${keyLine(device.windowsKeyEnds)}`.trim();
  }

  return { kicker, big: type, tone, toneKind, lines: [say, ...facts], warning: null };
}

/** The sentence under the cards, read from the scan's own reason for a personal copy. */
const personalWarning = (device) => (device.officePersonal === true && device.officePersonalReason
  ? `${device.officePersonalReason}. Remove the personal one so work files stay on the company account.`
  : null);

function officeCard(device) {
  const kicker = 'Office';
  if (isOldScan(device)) return oldCard(kicker);

  const type = device.officeLicenseType ?? 'Unknown';

  if (type === 'None') {
    return {
      kicker, big: 'None', tone: 'Not installed', toneKind: 'unknown',
      lines: ['No Office on this machine. Free apps (Teams, OneNote) are not counted.'], warning: null,
    };
  }
  if (type === 'Unknown') {
    return {
      kicker, big: 'Unknown', tone: 'Not reported', toneKind: 'unknown',
      lines: ['This scan did not read the Office licence.'], warning: null,
    };
  }

  const version = device.officeVersion ? `Office ${device.officeVersion}` : null;

  if (type === 'Microsoft 365') {
    const signedIn = Array.isArray(device.officeSignedInLicences) ? device.officeSignedInLicences : [];
    const companyPresent = signedIn.some((id) => !CONSUMER.test(id));
    const personal = device.officePersonal === true;
    const edition = signedIn.some((id) => /business/i.test(id)) ? 'Business' : 'Microsoft 365';
    const facts = [[edition, device.officeChannel, device.officeVersion].filter(Boolean).join(' · ')];
    if (device.officeLicensedTo) facts.push(`Licensed to ${device.officeLicensedTo}`);
    const checks = dayLabel(device.officeLicenceValidUntil);
    if (checks) facts.push(`Checks in again by ${checks}`);

    if (personal && companyPresent) {
      return {
        kicker, big: 'Microsoft 365', tone: 'Personal copy beside it', toneKind: 'warn',
        lines: ['A Home or Personal subscription is installed next to the company one.', ...facts],
        warning: personalWarning(device),
      };
    }
    if (personal) {
      return {
        kicker, big: 'Microsoft 365', tone: 'Personal account', toneKind: 'bad',
        lines: ['Only a personal Microsoft account is signed in. Work files go to the wrong OneDrive.', ...facts],
        warning: personalWarning(device),
      };
    }
    return {
      kicker, big: 'Microsoft 365', tone: 'Company account', toneKind: 'good',
      lines: ['Signed in with a work or school account.', ...facts], warning: null,
    };
  }

  if (type === 'Perpetual (volume)') {
    return {
      kicker, big: 'Perpetual (volume)', tone: 'Licensed', toneKind: 'good',
      lines: ['Office bought once, as a company (volume) licence.', ...(version ? [version] : [])], warning: null,
    };
  }

  // Perpetual (retail): the one remaining type the scan can name.
  return {
    kicker, big: 'Perpetual (retail)', tone: 'Licensed', toneKind: 'good',
    lines: ['Office bought once, with a retail key.', ...(version ? [version] : [])], warning: null,
  };
}

function joinedCard(device) {
  const kicker = 'Joined to';
  const join = device.deviceJoin ?? 'Unknown';

  if (isOldScan(device) || join === 'Unknown') {
    return {
      kicker, big: 'Unknown', tone: 'Old scan or Windows 7/8', toneKind: 'unknown',
      lines: ['Windows before 10 cannot say; an old scan did not ask.'], warning: null,
    };
  }
  if (join === 'Not joined') {
    return {
      kicker, big: 'Nothing', tone: 'Not joined', toneKind: 'bad',
      lines: ['A local PC. Company sign-in and policies do not reach it.'], warning: null,
    };
  }

  const big = device.tenantName || 'Company';
  if (join === 'Registered only') {
    return {
      kicker, big, tone: 'Registered only', toneKind: 'warn',
      lines: ['A work account is added, but the PC itself is not joined.'], warning: null,
    };
  }
  if (join === 'Azure AD + Domain') {
    return { kicker, big, tone: 'Azure AD + domain', toneKind: 'good', lines: ['Joined both ways.'], warning: null };
  }
  if (join === 'Domain') {
    return {
      kicker, big, tone: 'Domain', toneKind: 'good',
      lines: ['Joined to the company domain; policies apply.'], warning: null,
    };
  }
  // Azure AD on its own.
  return {
    kicker, big, tone: 'Azure AD', toneKind: 'good',
    lines: ['Signs in with company accounts; policies apply.', 'Not on a domain'], warning: null,
  };
}

/** The three cards in the order the page shows them: Windows, Office, Joined to. */
export function licenceCards(device) {
  return [windowsCard(device), officeCard(device), joinedCard(device)];
}

/** The amber box under the cards: the first warning any card carries, or null. */
export function licenceWarning(device) {
  return licenceCards(device).map((card) => card.warning).find(Boolean) ?? null;
}
