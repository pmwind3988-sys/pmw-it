import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useLocation } from 'react-router-dom';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { SessionContext, SESSION_PHASES } from '../hooks/useSession';
import { isInteractionRequired, isSilentRenewalUnavailable } from '../utils/authErrors';
import { withTimeout } from '../utils/timeout';
import {
  FRESH_SIGNIN_KEY,
  REAUTH_REDIRECT_KEY,
  clearFlag,
  forgetUser,
  hasFlag,
  readLastUser,
  rememberUser,
  setFlag,
} from '../utils/sessionKeys';
import { loginRequest } from '../authConfig';
import SessionDialog from '../components/SessionDialog';
import SignInTransition from '../components/SignInTransition';

/**
 * Automatic re-sign-in, and the two pieces of screen furniture that make it
 * legible: the dialog that runs during a recovery, and the entrance animation
 * that plays after a successful one.
 *
 * What this is built to avoid is signing out someone who was never signed out.
 * Two rules keep that from happening:
 *
 *   1. Nothing here is on a clock. There is no idle timer and no expiry watcher.
 *      A recovery starts only when a token call comes back with proof that Azure
 *      AD wants the user in front of it again (`isInteractionRequired`), or when
 *      the account has vanished from the MSAL cache on a page that needs one.
 *
 *   2. A recovery always tries `ssoSilent` first. If the Azure AD session cookie
 *      is alive — the common case, where only the cached *token* aged out — it
 *      comes back in about a second, the dialog closes, and the caller retries
 *      as if nothing happened. Nobody leaves the page. Only a second refusal,
 *      which means the session itself is gone, escalates to a redirect.
 *
 * The redirect, when it is reached, is immediate and unannounced by design.
 * `navigateToLoginRequestUrl` is on, so Microsoft returns the user to the URL
 * they were on; unsaved work on that page does not survive, which is the price
 * of not making them click through a dialog to get back in.
 */

// Long enough for a hidden iframe on a slow connection, short enough that the
// dialog never feels stuck. Past this we treat the silent path as unavailable
// rather than waiting out MSAL's own much longer window — which the caller has
// usually just spent already before handing the problem over.
const SSO_TIMEOUT_MS = 8000;

// How long "Signed back in" stays up. Much under a second and a successful
// recovery reads as a glitch rather than as something that worked.
const RECOVERED_HOLD_MS = 1100;

// The entrance animation gives up waiting for the page's data at this point and
// fades out anyway. A slow SharePoint read must not hold a veil over the app.
const ENTER_MAX_MS = 4000;

/** A caller that must not proceed and must not error: the page is navigating. */
const NEVER = new Promise(() => {});

export function SessionProvider({ children }) {
  const { instance, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const { pathname } = useLocation();

  const [phase, setPhase] = useState(SESSION_PHASES.IDLE);

  // Who this device knows about, read once at mount — a sign-in always arrives
  // through a redirect, so there is always a fresh mount behind a new name.
  const [storedUser] = useState(readLastUser);
  // Set the moment sign-out is asked for, and never unset: between that click
  // and the browser actually leaving for Microsoft, the account disappears from
  // the cache, which otherwise looks exactly like a timeout to the rule below.
  const [signedOut, setSignedOut] = useState(false);

  // The flag is read but not cleared here: StrictMode mounts twice in
  // development, and a read-and-clear would spend the animation on the mount
  // that gets thrown away. It is cleared when the animation finishes instead.
  const [entering, setEntering] = useState(() => hasFlag(FRESH_SIGNIN_KEY));
  const [contentReady, setContentReady] = useState(false);

  const phaseRef = useRef(phase);
  const recoveryRef = useRef(null);
  const redirectingRef = useRef(false);

  const toPhase = useCallback((next) => {
    phaseRef.current = next;
    setPhase(next);
  }, []);

  const account = instance.getActiveAccount() || instance.getAllAccounts()[0] || null;
  const knownUser = signedOut ? '' : account?.username || storedUser;

  // Remember who is signed in while there is someone to remember. The record
  // outlives the MSAL account on purpose — it is the login hint a silent
  // recovery needs, and the difference between "timed out" and "never signed in
  // on this device".
  useEffect(() => {
    if (account?.username) rememberUser(account.username);
  }, [account?.username]);

  /**
   * One recovery at a time, shared by every caller that hit the wall at once —
   * the dashboard alone fires two token calls in parallel, and two dialogs (or
   * two redirects) would be a bug.
   */
  const recover = useCallback(() => {
    if (redirectingRef.current) return NEVER;
    if (recoveryRef.current) return recoveryRef.current;

    const run = (async () => {
      toPhase(SESSION_PHASES.RECOVERING);
      // Read at call time, not closed over, so `recover` keeps one identity for
      // the lifetime of the provider and the effect below cannot re-fire on it.
      const loginHint =
        instance.getActiveAccount()?.username ||
        instance.getAllAccounts()[0]?.username ||
        readLastUser();

      // Without a hint `ssoSilent` cannot even be attempted — MSAL rejects it
      // for want of an account to be silent about. Straight to the redirect.
      if (loginHint) {
        try {
          const result = await withTimeout(
            instance.ssoSilent({ ...loginRequest, loginHint }),
            SSO_TIMEOUT_MS,
            'Silent sign-in timed out'
          );
          if (result?.account) {
            instance.setActiveAccount(result.account);
            rememberUser(result.account.username);
          }
          clearFlag(REAUTH_REDIRECT_KEY);
          toPhase(SESSION_PHASES.RECOVERED);
          return true;
        } catch (error) {
          // Refused, or the silent channel is unusable in this browser: either
          // way the redirect is the remaining route. Anything else — a dead
          // network, an outage — is not a dead session, so the dialog goes away
          // and the caller's own error handling has the failure.
          if (!isInteractionRequired(error) && !isSilentRenewalUnavailable(error)) {
            toPhase(SESSION_PHASES.IDLE);
            throw error;
          }
          console.debug('[session] silent sign-in unavailable — falling back to redirect');
        }
      }

      // A second refusal in one tab means the redirect is not fixing anything.
      // Stop, and let the user drive.
      if (hasFlag(REAUTH_REDIRECT_KEY)) {
        toPhase(SESSION_PHASES.BLOCKED);
        return false;
      }

      setFlag(REAUTH_REDIRECT_KEY);
      redirectingRef.current = true;
      toPhase(SESSION_PHASES.REDIRECTING);
      try {
        await instance.loginRedirect({
          ...loginRequest,
          loginHint: loginHint || undefined,
        });
      } catch (error) {
        // The handoff itself failed (an interaction already in progress, a
        // blocked navigation). Nothing is in flight, so let the button decide.
        console.error('[session] re-sign-in redirect failed:', error);
        clearFlag(REAUTH_REDIRECT_KEY);
        redirectingRef.current = false;
        toPhase(SESSION_PHASES.BLOCKED);
      }
      return false;
    })();

    recoveryRef.current = run;
    run.catch(() => {}).finally(() => {
      recoveryRef.current = null;
    });
    return run;
  }, [instance, toPhase]);

  /**
   * The one way this app asks for a token. Silent first; a recovery only on
   * proof; and on the far side of a successful recovery, the same silent call
   * once more so the caller gets what it originally asked for.
   */
  const acquireToken = useCallback(
    async (request) => {
      const cached = instance.getActiveAccount() || instance.getAllAccounts()[0] || null;
      if (cached && !instance.getActiveAccount()) instance.setActiveAccount(cached);

      if (cached) {
        try {
          return await instance.acquireTokenSilent({ ...request, account: cached });
        } catch (error) {
          // Two ways in: Azure AD refused, or MSAL's own silent renewal never
          // came back. The second is the one users actually meet — it is what
          // put `timed_out: See https://aka.ms/msal.js.errors` on the dashboard
          // where the figures should be. Both get a recovery; neither redirects
          // on its own, because `ssoSilent` still has to fail first.
          if (!isInteractionRequired(error) && !isSilentRenewalUnavailable(error)) throw error;
        }
      }

      const recovered = await recover();
      // Redirecting or blocked: the dialog owns the screen now. Hanging here is
      // deliberate — resolving would flash a half-built page and rejecting would
      // flash an error banner, both underneath a modal that already says what is
      // happening.
      if (!recovered) return NEVER;

      const account = instance.getActiveAccount() || instance.getAllAccounts()[0] || null;
      if (!account) return NEVER;
      return instance.acquireTokenSilent({ ...request, account });
    },
    [instance, recover]
  );

  const signOut = useCallback(() => {
    // Before the redirect, not after: leaving the hint behind would let the
    // auto-sign-in read a deliberate sign-out as a timeout and undo it.
    forgetUser();
    setSignedOut(true);
    clearFlag(REAUTH_REDIRECT_KEY);
    clearFlag(FRESH_SIGNIN_KEY);
    return instance.logoutRedirect({
      postLogoutRedirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:5173',
    });
  }, [instance]);

  const signIn = useCallback(() => {
    clearFlag(REAUTH_REDIRECT_KEY);
    redirectingRef.current = true;
    toPhase(SESSION_PHASES.REDIRECTING);
    return instance.loginRedirect({ ...loginRequest, loginHint: readLastUser() || undefined });
  }, [instance, toPhase]);

  /**
   * Someone known to this device is on a page that needs an account and hasn't
   * got one — the shape a timeout takes when it is noticed between calls rather
   * than during one. `/login` is exempt: that screen is a deliberate stop, and
   * signing people in as they arrive at it would take the door off its hinges.
   */
  const eligibleForAutoSignIn =
    !isAuthenticated && !!knownUser && pathname !== '/login' && !hasFlag(REAUTH_REDIRECT_KEY);

  useEffect(() => {
    if (!eligibleForAutoSignIn) return;
    if (inProgress !== InteractionStatus.None) return;
    if (phaseRef.current !== SESSION_PHASES.IDLE) return;
    recover().catch(() => {});
  }, [eligibleForAutoSignIn, inProgress, recover]);

  // "Signed back in" is a beat, not a screen.
  useEffect(() => {
    if (phase !== SESSION_PHASES.RECOVERED) return undefined;
    const id = setTimeout(() => toPhase(SESSION_PHASES.IDLE), RECOVERED_HOLD_MS);
    return () => clearTimeout(id);
  }, [phase, toPhase]);

  // The animation waits for the page to say it has its data, but never for long.
  useEffect(() => {
    if (!entering) return undefined;
    const id = setTimeout(() => setContentReady(true), ENTER_MAX_MS);
    return () => clearTimeout(id);
  }, [entering]);

  const markContentReady = useCallback(() => setContentReady(true), []);

  const endTransition = useCallback(() => {
    clearFlag(FRESH_SIGNIN_KEY);
    setEntering(false);
  }, []);

  const value = useMemo(
    () => ({
      phase,
      acquireToken,
      signIn,
      signOut,
      markContentReady,
      // What the screens ask before sending anyone to `/login`: a recovery is
      // running, or is about to be, so hold still rather than announcing that
      // the user is signed out.
      recovering: phase !== SESSION_PHASES.IDLE || eligibleForAutoSignIn,
    }),
    [phase, acquireToken, signIn, signOut, markContentReady, eligibleForAutoSignIn]
  );

  return (
    <SessionContext.Provider value={value}>
      {children}
      <SessionDialog phase={phase} onSignIn={signIn} />
      {entering && (
        <SignInTransition
          name={account?.name || ''}
          pathname={pathname}
          ready={contentReady}
          onDone={endTransition}
        />
      )}
    </SessionContext.Provider>
  );
}
