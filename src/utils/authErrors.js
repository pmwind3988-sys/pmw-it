import { InteractionRequiredAuthError } from '@azure/msal-browser';

/**
 * The one place that decides what counts as "the session is gone".
 *
 * This is the guard rail on automatic sign-in. Only an error that proves Azure
 * AD wants the user in front of it again may start a re-sign-in; a network
 * outage, one of this app's own 30s timeouts, a SharePoint 500 or a throttle
 * must stay an ordinary error and reach the screen as one. Signing someone back
 * in who never lost their session is a worse failure than the error they were
 * about to be shown, so anything ambiguous belongs on the "not a timeout" side
 * of this function — `monitor_window_timeout` in particular, which is usually a
 * blocked third-party cookie rather than an expired session.
 */
const INTERACTION_ERROR_CODES = new Set([
  'login_required',
  'interaction_required',
  'consent_required',
  'user_login_error',
  'no_account_error',
  'no_account_in_silent_request',
  // AAD's answer when the refresh token behind a silent call is dead. MSAL
  // usually wraps it, but not on every path.
  'invalid_grant',
]);

/**
 * AADSTS codes that say the same thing in the message body when the error code
 * itself is generic: no session (50058), the token is past its life (50173,
 * 700084), or the account needs to re-authenticate (50078).
 */
const INTERACTION_MESSAGE_CODES = /AADSTS(50058|50078|50173|700084)/;

/**
 * The other way a timed-out session shows up here, and in practice the common
 * one: MSAL renews silently through a hidden iframe to login.microsoftonline.com,
 * and when that iframe cannot complete — third-party cookies blocked, the frame
 * refused, the round trip too slow — it gives up with `timed_out` rather than
 * with anything AAD said. The page is left holding a raw MSAL error where its
 * data should be.
 *
 * This is kept apart from `isInteractionRequired` because it calls for a
 * different move: re-running `ssoSilent` would go down the same blocked channel
 * and fail the same way, so a recovery started from here skips the silent step
 * and asks for a redirect — the one route that does not depend on the iframe.
 */
const SILENT_CHANNEL_ERROR_CODES = new Set([
  'timed_out',
  'monitor_window_timeout',
  'silent_prompt_value_error',
]);

export function isSilentRenewalUnavailable(error) {
  if (!error) return false;
  const code = error.errorCode || error.code || '';
  if (SILENT_CHANNEL_ERROR_CODES.has(code)) return true;
  // This app's own wrapper (`withTimeout`) rejects with a plain Error.
  return /timed out/i.test(String(error.message || ''));
}

export function isInteractionRequired(error) {
  if (!error) return false;
  if (error instanceof InteractionRequiredAuthError) return true;

  const code = error.errorCode || error.code || '';
  if (INTERACTION_ERROR_CODES.has(code)) return true;

  const message = String(error.errorMessage || error.message || '');
  return INTERACTION_MESSAGE_CODES.test(message);
}
