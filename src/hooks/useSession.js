import { createContext, useContext } from 'react';

/**
 * Where the sign-in guard lives, kept apart from the provider that fills it:
 * a file exporting a hook next to a component drops out of Fast Refresh, and
 * eslint fails the build over it (same reason `initialsOf` has its own file).
 */

/**
 * `idle`        nothing is wrong — the normal state, and the only one that
 *               draws no dialog.
 * `recovering`  a silent re-sign-in is in flight. If the Azure AD session is
 *               still alive this is the only phase the user ever sees.
 * `recovered`   it worked. Held on screen briefly so the recovery is legible
 *               rather than a flicker, then back to idle.
 * `redirecting` the session really is gone; the browser is on its way to
 *               Microsoft.
 * `blocked`     we came back from that redirect and it still failed. The dialog
 *               stops trying and offers a button, because a third attempt is a
 *               loop, not a fix.
 */
export const SESSION_PHASES = Object.freeze({
  IDLE: 'idle',
  RECOVERING: 'recovering',
  RECOVERED: 'recovered',
  REDIRECTING: 'redirecting',
  BLOCKED: 'blocked',
});

export const SessionContext = createContext(null);

export function useSession() {
  const context = useContext(SessionContext);
  if (!context) {
    throw new Error('useSession must be used within SessionProvider');
  }
  return context;
}
