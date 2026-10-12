/**
 * Who is emailed when a shared checklist is signed, and whether anybody is.
 *
 * Held as one JSON value in the `IT Portal Settings` list, under
 * `EMAIL_SETTING_KEY`. The portal writes it (as whoever SharePoint lets edit
 * that list) and the server reads it at the moment a checklist is signed —
 * and again on every resend, so a resend can only ever go to the group IT
 * set here, never to an address the browser asked for.
 *
 * Pure, and shared by both sides, so the two cannot disagree about what a
 * stored value means.
 */

export const SETTINGS_LIST_NAME = 'IT Portal Settings';
export const EMAIL_SETTING_KEY = 'checklistEmail';

// At most this many recipients. A distribution group is ONE address; this is
// a list of groups and people, not a mailing list.
export const MAX_RECIPIENTS = 10;

// Deliberately plain: one @, something either side, a dot in the domain, no
// spaces. Exchange is the real judge; this only catches a typo before saving.
const ADDRESS = /^[^\s@<>()",;:]+@[^\s@<>()",;:]+\.[^\s@<>()",;:]+$/;

export const isEmailAddress = (value) => ADDRESS.test(String(value ?? '').trim());

export const DEFAULT_EMAIL_SETTINGS = Object.freeze({ enabled: false, to: [] });

/** A stored value (string or object) as `{ enabled, to }`; anything odd is off. */
export function normaliseEmailSettings(raw) {
  let value = raw;
  if (typeof raw === 'string') {
    try {
      value = JSON.parse(raw);
    } catch {
      return { ...DEFAULT_EMAIL_SETTINGS, to: [] };
    }
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ...DEFAULT_EMAIL_SETTINGS, to: [] };
  }
  const to = [];
  for (const entry of Array.isArray(value.to) ? value.to : []) {
    const address = String(entry ?? '').trim().toLowerCase();
    if (isEmailAddress(address) && !to.includes(address)) to.push(address);
  }
  return { enabled: value.enabled === true && to.length > 0, to: to.slice(0, MAX_RECIPIENTS) };
}

/**
 * One typed address checked against the ones already on the list. Answers
 * `{ address }` or `{ error }` in words for the person typing it.
 */
export function addRecipient(list, typed) {
  const address = String(typed ?? '').trim().toLowerCase();
  if (!address) return { error: 'Type an email address.' };
  if (!isEmailAddress(address)) return { error: `"${address}" is not an email address.` };
  if (list.includes(address)) return { error: `${address} is already on the list.` };
  if (list.length >= MAX_RECIPIENTS) return { error: `At most ${MAX_RECIPIENTS} addresses.` };
  return { address };
}

/** What is stored: always the normalised shape, so the server reads it as written. */
export const serialiseEmailSettings = (settings) => JSON.stringify(normaliseEmailSettings(settings));
