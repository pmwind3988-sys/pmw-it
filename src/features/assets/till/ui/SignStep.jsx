import TillSheet from './TillSheet';
import Button from '../../../../components/ui/Button';
import SignatureField from '../../ui/SignatureField';
import { Check } from '../../../../components/ui/Icons';

/**
 * Checkout for a handover: the person signs, like paying at a till. Asked
 * for, never insisted on — the laptop is in their hands either way.
 */
export default function SignStep({
  person, count, terms, rows, signature, onSignature, onConfirm, onClose, busy, busyLabel, failure,
}) {
  return (
    <TillSheet title={`${person.name} signs for ${count} item${count === 1 ? '' : 's'}`} onClose={onClose} wide>
      <p className="till-sheet-lede">{terms}</p>
      <ul className="till-sign-list">
        {rows.map((row) => (
          <li key={row.id}>
            <span>
              <strong>{row.name}</strong>
              {row.sub && <span className="till-mono">{row.sub}</span>}
            </span>
            <span className="till-mono">×{row.qty}</span>
          </li>
        ))}
      </ul>

      <SignatureField
        label={`${person.name.split(' ')[0]}’s signature`}
        value={signature}
        onChange={onSignature}
        disabled={busy}
      />

      {failure && <p className="till-sheet-error" role="alert">{failure}</p>}
      <div className="till-sign-actions">
        <Button icon={Check} className="till-cta" loading={busy} onClick={onConfirm}>
          {busy ? busyLabel : (signature ? 'Confirm handover' : 'Hand over without a signature')}
        </Button>
      </div>
    </TillSheet>
  );
}
