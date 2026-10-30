import React from 'react';
import ReactDOM from 'react-dom/client';
import PublicChecklist from './PublicChecklist';
import '../index.css';
import '../App.css';
import '../styles/shell.css';
import '../styles/forms.css';
import '../styles/public.css';

/**
 * The page a shared checklist link opens: `/c/<code>`, with no sign-in.
 *
 * A separate entry from `main.jsx` on purpose. It loads no MSAL, no router and
 * no portal page, so nothing on it leads anywhere else; every other address
 * still goes through the portal's sign-in.
 *
 * No theme toggle either, and always light, whatever the device is set to:
 * it is a form somebody signs and prints, and it should look like the paper
 * one on every phone.
 */
document.documentElement.setAttribute('data-theme', 'light');

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <PublicChecklist />
  </React.StrictMode>,
);
