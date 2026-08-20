import { useEffect, useRef, useState } from 'react';
import Logo from './Logo';

/**
 * The beat between "signed in" and "here is your dashboard".
 *
 * A successful sign-in returns from Microsoft into an empty shell that then
 * spends a second or two fetching the SharePoint list — long enough to read as
 * a stall on the first screen someone sees after handing over their password.
 * This covers that gap with something deliberate: the mark, a couple of rings
 * going out from it, and a line naming who just got signed in. It leaves by
 * dissolving into the page it was covering, so the dashboard appears to have
 * been what was behind it all along.
 *
 * It is decoration over content that is already there, so it is `aria-hidden`
 * with a single polite live line for anyone who is not looking at it, and it is
 * held nearly still under `prefers-reduced-motion` (see shell.css).
 */

// Below this the veil is a flash rather than a transition — worse than no veil.
// Measured from mount, so a fast cache hit still gets a deliberate-looking beat.
const MIN_VISIBLE_MS = 900;

// Matches the fade in shell.css. Unmounting earlier cuts the animation off.
const FADE_MS = 620;

export default function SignInTransition({ name, pathname, ready, onDone }) {
  const [leaving, setLeaving] = useState(false);
  const mountedAt = useRef(0);

  // Stamped in an effect rather than at `useRef(Date.now())`: the clock is a
  // side effect, and reading it during render is what the compiler objects to.
  // Declared first, so it is set before the effect below can read it.
  useEffect(() => {
    mountedAt.current = Date.now();
  }, []);

  useEffect(() => {
    if (!ready || leaving) return undefined;
    const remaining = Math.max(0, MIN_VISIBLE_MS - (Date.now() - mountedAt.current));
    const id = setTimeout(() => setLeaving(true), remaining);
    return () => clearTimeout(id);
  }, [ready, leaving]);

  useEffect(() => {
    if (!leaving) return undefined;
    const id = setTimeout(onDone, FADE_MS);
    return () => clearTimeout(id);
  }, [leaving, onDone]);

  const firstName = (name || '').trim().split(/\s+/)[0] || '';
  const heading = firstName ? `Welcome back, ${firstName}` : 'Welcome back';
  const sub =
    pathname === '/dashboard' || pathname === '/'
      ? 'Loading your dashboard…'
      : 'Getting your workspace ready…';

  return (
    <div className={`signin-veil${leaving ? ' leaving' : ''}`}>
      <div className="signin-glow" aria-hidden="true" />

      <div className="signin-core" aria-hidden="true">
        <span className="signin-ring" />
        <span className="signin-ring b" />
        <span className="signin-ring c" />
        <span className="signin-badge">
          <Logo size={44} />
        </span>
      </div>

      <p className="signin-title" aria-hidden="true">
        {heading}
      </p>
      <p className="signin-sub" aria-hidden="true">
        {sub}
      </p>

      <div className="signin-bar" aria-hidden="true">
        <span />
      </div>

      <p className="sr-only" role="status" aria-live="polite">
        {heading}. {sub}
      </p>
    </div>
  );
}
