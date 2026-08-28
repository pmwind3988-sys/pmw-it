import { useState } from 'react';
import SignatureDialog from '../../../components/SignatureDialog';
import Button from '../../../components/ui/Button';
import { Check, X, ClipboardList } from '../../../components/ui/Icons';

/**
 * "Sign here" — asked for, never insisted on.
 *
 * A signature is what turns "the register says Amir has the laptop" into
 * something anybody can stand behind six months later. It is not, however,
 * worth blocking a handover for: the laptop is in Amir's hands whether or not
 * a finger drew a squiggle on the phone, and a screen that refuses to record
 * that is a screen people work around by not recording anything.
 *
 * So it says recommended, and hands over without one if that is what happens.
 * What it will not do is imply somebody signed when they did not — a skipped
 * signature is stored as nothing at all.
 */
export default function SignatureField({
  label = 'Signature',
  hint = 'Recommended — ask the person to sign. It can be skipped.',
  value,
  onChange,
  disabled = false,
}) {
  const [signing, setSigning] = useState(false);

  return (
    <div className="as-signature">
      <div className="as-signature-head">
        <span className="as-field-label">
          {label}
          <span className="as-guess" title="Asked for, not required">recommended</span>
        </span>
        {!value && <span className="as-sub">{hint}</span>}
      </div>

      {value ? (
        <div className="as-signature-done">
          {/* The signature itself, not a tick claiming there is one. Somebody
              signing the wrong box has to be able to see that they did. */}
          <img src={value} alt={`${label} — signed`} className="as-signature-shot" />
          <span className="as-signature-ok"><Check size={13} /> Signed</span>
          <button
            type="button"
            className="as-iconbtn"
            onClick={() => onChange(null)}
            disabled={disabled}
            aria-label="Remove the signature and sign again"
          >
            <X size={13} />
          </button>
        </div>
      ) : (
        <Button
          variant="secondary"
          size="sm"
          icon={ClipboardList}
          disabled={disabled}
          onClick={() => setSigning(true)}
        >
          Sign here
        </Button>
      )}

      {signing && (
        <SignatureDialog
          onClose={() => setSigning(false)}
          onSave={(dataUrl) => {
            // `null` comes back when the pad was closed without a mark on it,
            // which is a skip and must not be stored as a signature.
            onChange(dataUrl || null);
            setSigning(false);
          }}
        />
      )}
    </div>
  );
}
