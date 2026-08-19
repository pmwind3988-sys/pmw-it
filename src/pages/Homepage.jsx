import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';

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

  useEffect(() => {
    document.title = 'PMW IT';
  }, []);

  useEffect(() => {
    // Wait until MSAL finishes any in-progress interaction (redirect handling,
    // silent refresh, MFA, etc.)
    if (inProgress !== InteractionStatus.None) return;
    window.location.replace(isAuthenticated ? '/dashboard' : '/login');
  }, [isAuthenticated, inProgress]);
  // No navigate in deps — window.location.replace can't trigger a re-render

  return (
    <div className="shell-gate">
      <div className="spinner" />
      <p>Opening the IT Service Portal…</p>
    </div>
  );
}
