import { useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from '../authConfig';
import Logo from '../components/Logo';
import IdleAnimation from '../components/IdleAnimation';

/** Microsoft's four-square mark. Their branding guidance wants it on the button that starts their sign-in. */
function MicrosoftMark() {
  return (
    <svg viewBox="0 0 20 20" width="18" height="18" aria-hidden="true" focusable="false" style={{ flexShrink: 0 }}>
      <rect x="0" y="0" width="9" height="9" fill="#F25022" />
      <rect x="11" y="0" width="9" height="9" fill="#7FBA00" />
      <rect x="0" y="11" width="9" height="9" fill="#00A4EF" />
      <rect x="11" y="11" width="9" height="9" fill="#FFB900" />
    </svg>
  );
}

/**
 * One door in, presented as a single centred card: the mark, the product name,
 * and the Microsoft button. Laid out as the OSHE portal's sign-in is — a poster
 * column beside a card column — so someone who uses both portals meets the same
 * screen twice.
 *
 * The idle animation is the poster column on a wide screen. A phone has no room
 * for a column, so rather than drop it the animation becomes a backdrop behind
 * the card: held to a wash so it reads as texture, with the card opaque over it.
 *
 * This is not a password form. Sign-in goes through the MSAL redirect.
 */
export default function LoginPage() {
  const { instance, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  useEffect(() => {
    document.title = 'PMW IT — Sign in';
  }, []);

  // Already signed in? The dashboard is where this screen was taking you.
  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;
    if (isAuthenticated) window.location.replace('/dashboard');
  }, [isAuthenticated, inProgress]);

  const busy = inProgress !== InteractionStatus.None;

  const handleLogin = () => instance.loginRedirect({ ...loginRequest, prompt: 'select_account' });

  return (
    <div className="auth-page">
      {/* Phone only: the same animation, dimmed to a backdrop rather than lost. */}
      <div className="auth-backdrop" aria-hidden="true">
        <IdleAnimation />
      </div>

      <div className="auth-poster">
        <div>
          <div className="auth-poster-name">PMW IT</div>
          <div className="auth-poster-kicker">Onboarding · Equipment · Access</div>
        </div>

        <IdleAnimation />

        <div className="auth-poster-foot">
          IT Service Portal · sign in with your Microsoft 365 work account
        </div>
      </div>

      <div className="auth-panelcol">
        <div className="auth-card rise">
          <Logo size={56} className="auth-logo" />

          <h1>IT Service Portal</h1>
          <p className="auth-sub">
            Sign in to raise onboarding and offboarding requests, track them, and hand over equipment.
          </p>

          <button type="button" className="auth-ms-btn" onClick={handleLogin} disabled={busy}>
            {busy ? <span className="auth-spinner" /> : <MicrosoftMark />}
            {busy ? 'Signing in…' : 'Continue with Microsoft 365'}
          </button>

          <hr className="auth-divider" />
          <p className="auth-note">Only PMW Microsoft 365 work accounts can sign in.</p>
        </div>
      </div>
    </div>
  );
}
