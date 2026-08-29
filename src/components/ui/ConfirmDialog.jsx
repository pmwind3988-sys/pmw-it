import { useEffect, useRef } from 'react';
import { AlertTriangle } from './Icons';

/**
 * "Are you sure" — asked properly.
 *
 * Everything that removes something asks first, because these screens are used
 * on a phone with one thumb and the difference between Save and Remove is a
 * few millimetres. `window.confirm` did the job on a desktop and is a poor fit
 * here: some phone browsers show it as an anonymous chrome bar with the site's
 * name on it, and none of them can say WHAT is about to go.
 *
 * The safe answer is the one already under the cursor. The dialog opens with
 * Cancel focused, Escape cancels, and Enter therefore cancels too — there is
 * no keystroke that destroys something without aiming at it first.
 */
export default function ConfirmDialog({
  title,
  body,
  confirmLabel = 'Remove',
  cancelLabel = 'Keep it',
  onAnswer,
}) {
  const cancelRef = useRef(null);

  useEffect(() => {
    cancelRef.current?.focus();
  }, []);

  useEffect(() => {
    const onKey = (event) => {
      if (event.key === 'Escape') onAnswer(false);
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onAnswer]);

  return (
    <div className="ui-confirm" role="dialog" aria-modal="true" aria-label={title}>
      {/* The backdrop is a real button rather than a click handler on a div, so
          that it is reachable and announced like the cancel it is. */}
      <button
        type="button"
        className="ui-confirm-back"
        aria-label="Cancel"
        onClick={() => onAnswer(false)}
      />
      <div className="ui-confirm-box">
        <h2 className="ui-confirm-title">
          <AlertTriangle size={16} /> {title}
        </h2>
        {body && <p className="ui-confirm-body">{body}</p>}
        <div className="ui-confirm-actions">
          <button
            type="button"
            className="ui-btn ui-btn-md ui-btn-secondary"
            ref={cancelRef}
            onClick={() => onAnswer(false)}
          >
            {cancelLabel}
          </button>
          <button
            type="button"
            className="ui-btn ui-btn-md ui-confirm-go"
            onClick={() => onAnswer(true)}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
