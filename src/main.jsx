import React from 'react';
import ReactDOM from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { BrowserRouter } from 'react-router-dom';
import { msalConfig } from './authConfig';
import { ThemeProvider } from './context/ThemeContext';
import { SessionProvider } from './context/SessionContext';
import { withTimeout } from './utils/timeout';
import {
  FRESH_SIGNIN_KEY,
  REAUTH_REDIRECT_KEY,
  clearFlag,
  rememberUser,
  setFlag,
} from './utils/sessionKeys';
import App from './App';
import './index.css';
import './App.css';
// Last, so the shell's tokens win over the pre-Shell defaults they replace.
import './styles/shell.css';
import './styles/auth.css';
import './styles/devices.css';

// MSAL request timeout (30 seconds)
const MSAL_TIMEOUT_MS = 30000;

async function bootstrap() {
  const msalInstance = new PublicClientApplication(msalConfig);

  await msalInstance.initialize();
  console.log('[MSAL] Instance initialized');

  try {
    // Wrap with timeout to prevent hanging on slow network
    const response = await withTimeout(
      msalInstance.handleRedirectPromise(),
      MSAL_TIMEOUT_MS,
      'Login redirect timed out'
    );
    if (response) {
      msalInstance.setActiveAccount(response.account);
      console.log('[MSAL] Redirect login completed for:', response.account?.username);
      // A redirect that came back with an account is the end of any re-sign-in
      // that started one, so the loop guard is spent — and the entrance
      // animation is owed a play, which React is not yet running to be told.
      clearFlag(REAUTH_REDIRECT_KEY);
      setFlag(FRESH_SIGNIN_KEY);
      rememberUser(response.account?.username);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        rememberUser(accounts[0].username);
      }
    }
  } catch (error) {
    // Check for timeout errors first
    if (error?.message?.includes('timed out')) {
      console.error('[MSAL] Request timeout:', error.message);
      // Continue with anonymous rendering - user can still interact
    }
    // ✅ no_token_request_cache_error is harmless — it just means there was
    // no redirect in progress when the page loaded (normal first visit / refresh).
    // Log it silently and continue rendering.
    else if (error?.errorCode === 'no_token_request_cache_error') {
      console.debug('[MSAL] No redirect in progress (normal on fresh load)');
      // Still restore account from cache if available
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0]);
    } else {
      console.error('[MSAL] Initialization error:', error);
    }
  }

  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <ThemeProvider>
          <BrowserRouter>
            {/* Inside the router: the guard exempts `/login` from automatic
                sign-in, and the entrance animation reads the path for its
                second line. */}
            <SessionProvider>
              <App />
            </SessionProvider>
          </BrowserRouter>
        </ThemeProvider>
      </MsalProvider>
    </React.StrictMode>
  );
}

bootstrap();