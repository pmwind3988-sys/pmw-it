import { useState } from 'react';
import Button from '../../../../components/ui/Button';
import { useGraphToken } from '../../../../hooks/useRequests';
import { formatMYT } from '../../../../utils/malaysiaTime';
import { emailStatus } from '../emailLog';
import { requestChecklistEmail } from '../../sharepoint/checklistEmailApi';

const stamp = (at) => {
  const time = Date.parse(at);
  return Number.isFinite(time) ? formatMYT(time, 'datetime12') : String(at ?? '');
};

const describe = (entry) => {
  if (!entry.ok) return `failed: ${entry.error || 'not sent'}`;
  if (entry.kind === 'resent') return `resent by ${entry.by || 'someone in IT'}`;
  return 'sent on signing';
};

const HEADING = { sent: 'Emailed', failed: 'Email not sent', none: 'Not emailed' };
const PILL = { sent: 'Sent', failed: 'Failed', none: 'Not sent' };
const CTA = { sent: 'Resend email', failed: 'Try again', none: 'Send email' };

/** The email block in a signed link's detail panel: what went out, and a way to send it again. */
export default function LinkEmailStatus({ link, onChanged }) {
  const getGraphToken = useGraphToken();
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const { state, last, lastSent, log } = emailStatus(link);

  const send = async () => {
    setBusy(true);
    setNote('');
    setError('');
    try {
      const { accessToken } = await getGraphToken();
      const result = await requestChecklistEmail({ graphToken: accessToken, action: 'resend', id: link.id });
      if (result.ok) {
        setNote(`Sent to ${result.to.join(', ')}.`);
        onChanged?.();
      } else {
        setError(result.error);
      }
    } catch (thrown) {
      setError(thrown.message || 'The email was not sent.');
    } finally {
      setBusy(false);
    }
  };

  let line;
  if (state === 'sent') {
    line = `${stamp(lastSent.at)} to ${lastSent.to.join(', ')}`;
  } else if (state === 'failed') {
    line = last.error || 'The last email did not go out.';
  } else {
    line = 'Signed before emails were turned on, or while they were off.';
  }

  return (
    <section className="cl-dsec cl-email" aria-label="Email">
      <div className="cl-email-head">
        <h3 className="cl-email-title">{HEADING[state]}</h3>
        <span className={`cl-pill cl-email-pill cl-email-pill-${state}`}>
          <i className="cl-dot" aria-hidden="true" />
          {PILL[state]}
        </span>
      </div>
      <p className="cl-email-line">{line}</p>

      <div className="cl-email-actions">
        <Button variant={state === 'sent' ? 'ghost' : 'primary'} loading={busy} onClick={send}>
          {CTA[state]}
        </Button>
      </div>
      {note && <p className="cl-email-ok" role="status">{note}</p>}
      {error && <p className="cl-email-err" role="alert">{error}</p>}

      {log.length > 0 && (
        <div className="cl-email-history">
          <h4 className="cl-dsec-h">History</h4>
          <ul>
            {[...log].reverse().map((entry, index) => (
              <li key={`${entry.at}-${index}`}>{stamp(entry.at)} · {describe(entry)}</li>
            ))}
          </ul>
        </div>
      )}
    </section>
  );
}
