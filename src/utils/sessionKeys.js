/**
 * The three keys the sign-in guard keeps outside React, and safe accessors for
 * them. Web storage throws outright in a locked-down browser (Safari private
 * mode, a hardened group policy), and a guard that crashes on its own bookkeeping
 * is worse than no guard at all — so every read and write here swallows the
 * failure and reports "nothing stored".
 *
 * LAST_USER_KEY (localStorage)
 *   Who signed in on this device. Kept deliberately after MSAL has evicted the
 *   account, because it is the only thing left to hand `ssoSilent` as a login
 *   hint — and because its absence is how we tell a timed-out session apart from
 *   someone who has simply never signed in here. Cleared on sign-out: leaving it
 *   would sign the user straight back in and make signing out impossible.
 *
 * REAUTH_REDIRECT_KEY (sessionStorage)
 *   A re-sign-in redirect is in flight. Survives the round trip to Microsoft and
 *   dies with the tab, which is exactly the lifetime needed to stop a sign-in
 *   that keeps failing from bouncing the browser in a loop.
 *
 * FRESH_SIGNIN_KEY (sessionStorage)
 *   The redirect came back with an account, so the entrance animation is owed
 *   one play. Set in the MSAL bootstrap, before React exists to be told.
 */

export const LAST_USER_KEY = 'pmw:last-user';
export const REAUTH_REDIRECT_KEY = 'pmw:reauth-redirect';
export const FRESH_SIGNIN_KEY = 'pmw:fresh-signin';

export function readLastUser() {
  try {
    return localStorage.getItem(LAST_USER_KEY) || '';
  } catch {
    return '';
  }
}

export function rememberUser(username) {
  if (!username) return;
  try {
    localStorage.setItem(LAST_USER_KEY, username);
  } catch {
    /* storage unavailable — silent recovery just loses its hint */
  }
}

export function forgetUser() {
  try {
    localStorage.removeItem(LAST_USER_KEY);
  } catch {
    /* nothing to do */
  }
}

export function hasFlag(key) {
  try {
    return sessionStorage.getItem(key) === '1';
  } catch {
    return false;
  }
}

export function setFlag(key) {
  try {
    sessionStorage.setItem(key, '1');
  } catch {
    /* nothing to do */
  }
}

export function clearFlag(key) {
  try {
    sessionStorage.removeItem(key);
  } catch {
    /* nothing to do */
  }
}
