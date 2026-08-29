import { useState } from 'react';
import Button from '../../../components/ui/Button';
import { Check, X, AlertTriangle } from '../../../components/ui/Icons';
import {
  personEditRefusal, personAt, normaliseEmail,
} from '../handover/planPersonEdit';

/**
 * Fixing a person's name and email in place.
 *
 * People arrive in this register typed at a desk in a hurry, or from a
 * directory that was unreachable at the time — so "amir" against
 * `amir@pmw.com` when the company is `pmwgroup.com` is a normal Tuesday. Until
 * now the only way to correct that was to return everything and hand it all
 * out again, which rewrites the dates, loses the signatures, and puts a return
 * in the record that never happened.
 *
 * So the person is editable and what they hold is not. Nothing on this form
 * can move an item, and the screen says so, because somebody about to retype
 * an email against a page listing four laptops deserves to know which of the
 * two things they are changing.
 */
export default function PersonEditor({
  current, handovers, busy = false, onSave, onCancel,
}) {
  const [draft, setDraft] = useState({
    name: current.name ?? '',
    email: current.email ?? '',
  });

  const refusal = personEditRefusal(draft, current);

  // Somebody else already at the new address. Not refused — it is usually the
  // point, since a typo is exactly what creates a second Amir — but far too
  // big a thing to happen without being said out loud first.
  const other = normaliseEmail(draft.email) === normaliseEmail(current.email)
    ? null
    : personAt(handovers, draft.email);

  const set = (field) => (event) => setDraft({ ...draft, [field]: event.target.value });

  return (
    <div className="as-personedit">
      <div className="as-form">
        <label className="as-field">
          <span className="as-field-label">Name</span>
          <input value={draft.name} onChange={set('name')} disabled={busy} autoComplete="off" />
        </label>
        <label className="as-field">
          <span className="as-field-label">Work email</span>
          <input
            type="email"
            value={draft.email}
            onChange={set('email')}
            disabled={busy}
            autoComplete="off"
            placeholder="amir@pmwgroup.com"
          />
        </label>
      </div>

      <p className="as-hint as-hint-inline">
        Only who they are. Everything they hold stays exactly as it is — the
        same items, the same dates, the same signatures — and their past
        handovers come with them rather than being left behind under the old
        address.
      </p>

      {other && (
        <p className="as-field-issue">
          <AlertTriangle size={13} />{' '}
          {other.name} already uses {other.email}, with {other.rows} handover
          {other.rows === 1 ? '' : 's'} on record. Saving joins the two into one person.
        </p>
      )}

      <div className="as-actions">
        <Button icon={Check} disabled={busy || Boolean(refusal)} onClick={() => onSave(draft)}>
          {busy ? 'Saving…' : 'Save'}
        </Button>
        <Button variant="ghost" icon={X} disabled={busy} onClick={onCancel}>
          Cancel
        </Button>
        {/* The reason the button is off, next to the button. A disabled
            control with no explanation is a dead end somebody presses twice
            and then walks away from. */}
        {refusal && <span className="as-sub">{refusal}</span>}
      </div>
    </div>
  );
}
