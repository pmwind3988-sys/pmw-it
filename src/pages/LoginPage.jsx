import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from '../authConfig';

export default function LoginPage() {
  const { instance, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  useEffect(() => {
    document.title = 'PMW IT - Sign In';
  }, []);

  // If already signed in, skip straight to the form
  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;
    if (isAuthenticated) {
      window.location.replace('/list');
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
          <svg width="20" height="20" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
            <rect x="1" y="1" width="9" height="9" fill="#F25022"/>
            <rect x="11" y="1" width="9" height="9" fill="#7FBA00"/>
            <rect x="1" y="11" width="9" height="9" fill="#00A4EF"/>
            <rect x="11" y="11" width="9" height="9" fill="#FFB900"/>
          </svg>
          Sign in with Microsoft
        </button>
      </div>
    </div>
  );
}