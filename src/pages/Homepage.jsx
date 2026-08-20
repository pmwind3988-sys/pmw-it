import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { useSession } from '../hooks/useSession';

/**
 * The entry point, and nothing else: it decides between the sign-in screen and
 * the dashboard, then gets out of the way.
 *
 * Uses window.location.replace() instead of React Router navigate().
 *
 * WHY: navigate() causes a React Router state update → re-render → effect runs
 * again → infinite loop. window.location.replace() is a real browser navigation
 * that happens entirely outside React's render cycle.
 */
export default function Homepage() {
  const { inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const { recovering } = useSession();

  useEffect(() => {
    document.title = 'PMW IT';
  }, []);

  useEffect(() => {
    // Wait until MSAL finishes any in-progress interaction (redirect handling,
    // silent refresh, MFA, etc.)
    if (inProgress !== InteractionStatus.None) return;
    // A session that timed out overnight is noticed here first. Leaving for
    // `/login` would be a real browser navigation, killing the silent re-sign-in
    // mid-flight and handing a button to someone the guard could have let
    // straight through.
    if (!isAuthenticated && recovering) return;
    window.location.replace(isAuthenticated ? '/dashboard' : '/login');
  }, [isAuthenticated, inProgress, recovering]);
  // No navigate in deps — window.location.replace can't trigger a re-render

  return (
    <div className="shell-gate">
      <div className="spinner" />
      <p>{recovering && !isAuthenticated ? 'Restoring your session…' : 'Opening the IT Service Portal…'}</p>
    </div>
  );
}
