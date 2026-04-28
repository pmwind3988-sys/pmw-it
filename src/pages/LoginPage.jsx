import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from '../authConfig';

export default function LoginPage() {
  const { instance, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  // If already signed in, skip straight to the form
  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;
    if (isAuthenticated) {
      window.location.replace('/it-boarding-form');
    }
  }, [isAuthenticated, inProgress]);

  const handleLogin = () => {
    instance.loginRedirect({ ...loginRequest, prompt: 'select_account' });
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <h1>IT Onboarding</h1>
        <p>Sign in with your Microsoft 365 account to continue</p>
        <button className="ms-button" onClick={handleLogin}
          disabled={inProgress !== InteractionStatus.None}>
          <svg viewBox="0 0 21 21" fill="currentColor">
            <path d="M11.6 0h-3.2v10.4h3.2V0zm-1.6 10.4c0-.9.7-1.5.1.6-1.5.9 0 1.6.6 1.6 1.5 0 .9-.7 1.6-1.6 1.6-.9.1-1.6-.7-1.6-1.6zm7.8-5.2H6.4v5.2h3.2V10h1.6V5.2h1.6V0h-1.6zM3.2 0H0v10.4h3.2V0z"/>
          </svg>
          Sign in with Microsoft
        </button>
      </div>
    </div>
  );
}