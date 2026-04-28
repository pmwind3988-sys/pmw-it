import React from 'react';
import ReactDOM from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { BrowserRouter } from 'react-router-dom';
import { msalConfig } from './authConfig';
import { ThemeProvider } from './context/ThemeContext';
import App from './App';
import './index.css';
import './App.css';

async function bootstrap() {
  const msalInstance = new PublicClientApplication(msalConfig);

  await msalInstance.initialize();
  console.log('[MSAL] Instance initialized');

  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response) {
      msalInstance.setActiveAccount(response.account);
      console.log('[MSAL] Redirect login completed for:', response.account?.username);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0]);
    }
  } catch (error) {
    // ✅ no_token_request_cache_error is harmless — it just means there was
    // no redirect in progress when the page loaded (normal first visit / refresh).
    // Log it silently and continue rendering.
    if (error?.errorCode === 'no_token_request_cache_error') {
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
            <App />
          </BrowserRouter>
        </ThemeProvider>
      </MsalProvider>
    </React.StrictMode>
  );
}

bootstrap();