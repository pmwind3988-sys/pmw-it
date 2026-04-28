import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';

/**
 * Homepage.jsx
 *
 * Uses window.location.replace() instead of React Router navigate().
 *
 * WHY: navigate() causes a React Router state update → re-render → effect
 * runs again → infinite loop. window.location.replace() is a real browser
 * navigation that happens entirely outside React's render cycle.
 */

/**
 * Homepage.jsx
 *
 * Uses window.location.replace() instead of React Router navigate().
 *
 * WHY: navigate() causes a React Router state update → re-render → effect
 * runs again → infinite loop. window.location.replace() is a real browser
 * navigation that happens entirely outside React's render cycle.
 */
export default function Homepage() {
  const { inProgress }  = useMsal();
  const isAuthenticated = useIsAuthenticated();

  useEffect(() => {
    document.title = 'PMW IT';
  }, []);

  useEffect(() => {
    // Wait until MSAL finishes any in-progress interaction
    // (redirect handling, silent refresh, MFA, etc.)
    if (inProgress !== InteractionStatus.None) return;

    if (isAuthenticated) {
      window.location.replace('/list');
    } else {
      window.location.replace('/login');
    }
  }, [isAuthenticated, inProgress]);
  // No navigate in deps — window.location.replace can't trigger a re-render

  return (
    <div className="home-page">
      <div className="login-card">
        <h1>IT Onboarding Portal</h1>
        <p>Redirecting…</p>
      </div>
    </div>
  );
}