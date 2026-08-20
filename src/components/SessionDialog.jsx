import { useEffect, useRef } from 'react';
import { SESSION_PHASES } from '../hooks/useSession';
import Logo from './Logo';
import { Check, ShieldCheck } from './ui/Icons';

/**
 * What a timed-out session looks like while it is being fixed.
 *
 * Without this the recovery is invisible: the screen simply stops answering for
 * a second or two, and then either carries on or jumps to Microsoft. The dialog
 * is there to name what is happening in the gap, so neither outcome arrives
 * unexplained.
 *
 * It is deliberately not dismissable while work is in flight — there is nothing
 * behind it that can be used until the token is back, and a closable dialog
 * would only invite people to close it and find the page still broken. The one
 * phase that takes an action, `blocked`, is the one phase with a button.
 */

const COPY = {
  [SESSION_PHASES.RECOVERING]: {
    title: 'Session timed out',
    body: 'Signing you back in — this usually takes a moment.',
  },
  [SESSION_PHASES.RECOVERED]: {
    title: 'Signed back in',
    body: 'Picking up where you left off.',
  },
  [SESSION_PHASES.REDIRECTING]: {
    title: 'Signing you in again',
    body: 'Taking you to Microsoft to confirm it is you.',
  },
  [SESSION_PHASES.BLOCKED]: {
    title: "We couldn't sign you back in",
    body: 'Your session has expired and the automatic retry did not work. Sign in again to carry on.',
  },
};

export default function SessionDialog({ phase, onSignIn }) {
  const buttonRef = useRef(null);

  // The button is the only thing left to do at this point, so put the keyboard
  // on it rather than leaving focus somewhere behind the scrim.
  useEffect(() => {
    if (phase === SESSION_PHASES.BLOCKED) buttonRef.current?.focus();
  }, [phase]);

  if (phase === SESSION_PHASES.IDLE) return null;

  const copy = COPY[phase];
  if (!copy) return null;

  const done = phase === SESSION_PHASES.RECOVERED;
  const blocked = phase === SESSION_PHASES.BLOCKED;

  return (
    <div className={`session-scrim${done ? ' out' : ''}`}>
      <div
        className="session-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="session-dialog-title"
        aria-describedby="session-dialog-body"
      >
        <div className="session-mark" aria-hidden="true">
          {done ? (
            <span className="session-tick">
              <Check size={22} />
            </span>
          ) : blocked ? (
            <span className="session-warn">
              <ShieldCheck size={22} />
            </span>
          ) : (
            <>
              <span className="session-pulse" />
              <span className="session-pulse b" />
              <span className="session-orbit" />
              <Logo size={30} />
            </>
          )}
        </div>

        <h2 id="session-dialog-title">{copy.title}</h2>
        <p id="session-dialog-body">{copy.body}</p>

        {blocked && (
          <button type="button" ref={buttonRef} className="session-btn" onClick={onSignIn}>
            Sign in with Microsoft
          </button>
        )}
      </div>
    </div>
  );
}
