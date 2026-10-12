import { useEffect, useRef, useState } from 'react';
import Button from '../../../../components/ui/Button';
import Spinner from '../../../../components/ui/Spinner';
import { useGraphToken, useSharePointToken } from '../../../../hooks/useRequests';
import { readEmailSettings, saveEmailSettings } from '../../sharepoint/emailSettings';
import { requestChecklistEmail } from '../../sharepoint/checklistEmailApi';
import { addRecipient } from '../emailSettings';

/**
 * Who is emailed when a shared checklist is signed. Saved for everyone in IT,
 * so it lives on the server's side of the page, not in this browser.
 */
export default function EmailSettingsDialog({ siteUrl, onClose }) {
  const getSpToken = useSharePointToken();
  const getGraphToken = useGraphToken();
  const firstControl = useRef(null);

  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [saved, setSaved] = useState({ enabled: false, to: [] });
  const [enabled, setEnabled] = useState(false);
  const [to, setTo] = useState([]);
  const [typed, setTyped] = useState('');
  const [typeError, setTypeError] = useState('');
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState('');
  const [saveNote, setSaveNote] = useState('');
  const [testing, setTesting] = useState(false);
  const [testError, setTestError] = useState('');
  const [testNote, setTestNote] = useState('');

  // Loaded when the dialog opens. Only the awaited result touches state.
  useEffect(() => {
    let live = true;
    (async () => {
      try {
        const { accessToken } = await getSpToken();
        const settings = await readEmailSettings({ siteUrl, token: accessToken });
        if (!live) return;
        setSaved(settings);
        setEnabled(settings.enabled);
        setTo(settings.to);
      } catch (thrown) {
        if (live) setLoadError(thrown.message || 'The email settings could not be loaded.');
      } finally {
        if (live) setLoading(false);
      }
    })();
    return () => { live = false; };
  }, [siteUrl, getSpToken]);

  useEffect(() => {
    if (!loading) firstControl.current?.focus();
  }, [loading]);

  // Captured on the window so Escape closes the dialog and goes no further,
  // whatever else on the page is listening for it.
  useEffect(() => {
    const onKey = (event) => {
      if (event.key !== 'Escape') return;
      event.stopImmediatePropagation();
      event.stopPropagation();
      onClose();
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [onClose]);

  const dirty = enabled !== saved.enabled || to.join('|') !== saved.to.join('|');
  const testReady = !dirty && saved.enabled === true;

  const add = () => {
    const result = addRecipient(to, typed);
    if (result.error) {
      setTypeError(result.error);
      return;
    }
    setTo([...to, result.address]);
    setTyped('');
    setTypeError('');
  };

  const onTypeKey = (event) => {
    if (event.key === 'Enter' || event.key === ',') {
      event.preventDefault();
      add();
    }
  };

  const remove = (address) => setTo(to.filter((entry) => entry !== address));

  const save = async () => {
    setSaving(true);
    setSaveError('');
    setSaveNote('');
    try {
      const { accessToken } = await getSpToken();
      const next = await saveEmailSettings({ siteUrl, token: accessToken, settings: { enabled, to } });
      setSaved(next);
      setEnabled(next.enabled);
      setTo(next.to);
      setSaveNote('Saved.');
    } catch (thrown) {
      setSaveError(thrown.message || 'The settings could not be saved. Your changes are still here.');
    } finally {
      setSaving(false);
    }
  };

  const sendTest = async () => {
    setTesting(true);
    setTestError('');
    setTestNote('');
    try {
      const { accessToken } = await getGraphToken();
      const result = await requestChecklistEmail({ graphToken: accessToken, action: 'test' });
      if (result.ok) setTestNote(`Test sent to ${result.to.join(', ')}.`);
      else setTestError(result.error);
    } catch (thrown) {
      setTestError(thrown.message || 'The test email was not sent.');
    } finally {
      setTesting(false);
    }
  };

  return (
    <div className="ui-confirm" role="dialog" aria-modal="true" aria-labelledby="cl-email-title">
      <button type="button" className="ui-confirm-back" aria-label="Close" onClick={onClose} />
      <div className="ui-confirm-box cl-email-dialog">
        <h2 id="cl-email-title" className="cl-email-title">Email when a checklist is signed</h2>

        {loading && (
          <p className="cl-email-lead"><Spinner /> Loading the settings…</p>
        )}

        {loadError && <p className="cl-email-err" role="alert">{loadError}</p>}

        {!loading && !loadError && (
          <>
            <label className={`ff-switch${enabled ? ' ff-switch-on' : ''}`}>
              <input
                ref={firstControl}
                type="checkbox"
                className="ff-sr-input"
                checked={enabled}
                onChange={(event) => setEnabled(event.target.checked)}
              />
              <span className="ff-switch-track" aria-hidden="true" />
              Send an email for each new signing
            </label>

            <div className="cl-email-field">
              <span className="cl-email-label" id="cl-email-to-label">Send to</span>
              {to.length > 0 && (
                <ul className="cl-email-chips" aria-labelledby="cl-email-to-label">
                  {to.map((address) => (
                    <li key={address} className="cl-email-chip">
                      <span>{address}</span>
                      <button
                        type="button"
                        className="cl-email-remove"
                        aria-label={`Remove ${address}`}
                        onClick={() => remove(address)}
                      >
                        ×
                      </button>
                    </li>
                  ))}
                </ul>
              )}
              <div className="cl-email-add">
                <input
                  type="text"
                  inputMode="email"
                  autoComplete="off"
                  className="cl-email-input"
                  placeholder="Add an address or a group"
                  aria-label="Add an address or a group"
                  aria-invalid={typeError ? true : undefined}
                  value={typed}
                  onChange={(event) => { setTyped(event.target.value); setTypeError(''); }}
                  onKeyDown={onTypeKey}
                />
                <Button variant="ghost" onClick={add}>Add</Button>
              </div>
              {typeError && <p className="cl-email-err" role="alert">{typeError}</p>}
              <p className="cl-email-help">
                A distribution group counts as one address. Saved for everyone in IT, not just this browser.
              </p>
            </div>

            <p className="cl-email-note">
              Only checklists signed from now on are emailed. Checklists already signed are left alone — send any of them by hand from its panel.
            </p>

            {enabled && to.length === 0 && (
              <p className="cl-email-warn">Add at least one address, or emails stay off.</p>
            )}

            {saveError && <p className="cl-email-err" role="alert">{saveError}</p>}
            {saveNote && <p className="cl-email-ok" role="status">{saveNote}</p>}
            {testError && <p className="cl-email-err" role="alert">{testError}</p>}
            {testNote && <p className="cl-email-ok" role="status">{testNote}</p>}
            {dirty && saved.enabled && (
              <p className="cl-email-help">Save your changes before sending a test email.</p>
            )}

            <div className="cl-email-actions">
              <Button loading={saving} onClick={save}>Save</Button>
              <Button variant="ghost" loading={testing} disabled={!testReady} onClick={sendTest}>
                Send a test email
              </Button>
              <button type="button" className="ui-btn ui-btn-md ui-btn-subtle" onClick={onClose}>
                Close
              </button>
            </div>
          </>
        )}

        {loadError && (
          <div className="cl-email-actions">
            <button type="button" className="ui-btn ui-btn-md ui-btn-subtle" onClick={onClose}>
              Close
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
