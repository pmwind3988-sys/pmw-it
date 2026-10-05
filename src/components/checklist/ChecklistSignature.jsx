import { useState } from 'react';
import Field from '../form/Field';
import Button from '../ui/Button';
import { Pencil } from '../ui/Icons';
import SignatureDialog from '../SignatureDialog';

/**
 * The signature on an asset checklist — required, unlike the handover pages'
 * `SignatureField`: here the signature IS the record.
 */
export default function ChecklistSignature({
  value, onChange, error, label = 'Your Signature', help = 'Sign in the middle of the box.',
}) {
  const [signing, setSigning] = useState(false);

  return (
    <Field label={label} required error={error} help={help} wide>
      {value ? (
        <div className="ff-signed">
          <img src={value} alt="Your signature" />
          <Button variant="ghost" size="sm" icon={Pencil} onClick={() => setSigning(true)}>
            Sign again
          </Button>
        </div>
      ) : (
        <button type="button" className="ff-signbtn" onClick={() => setSigning(true)}>
          <Pencil size={16} /> Tap to sign
        </button>
      )}

      {signing && (
        <SignatureDialog
          onSave={(dataUrl) => {
            onChange(dataUrl || null);
            setSigning(false);
          }}
          onClose={() => setSigning(false)}
        />
      )}
    </Field>
  );
}
