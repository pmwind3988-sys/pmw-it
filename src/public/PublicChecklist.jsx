import { useEffect, useState } from 'react';
import Logo from '../components/Logo';
import Button from '../components/ui/Button';
import { Card, ErrorBanner } from '../components/ui/Surfaces';
import { AlertTriangle, Check, Printer, Lock } from '../components/ui/Icons';
import ChecklistFields from '../components/checklist/ChecklistFields';
import ChecklistSignature from '../components/checklist/ChecklistSignature';
import ChecklistRecord from '../components/checklist/ChecklistRecord';
import { FORM_MODES, newItemRow } from '../features/forms/checklistForm';
import { validateChecklist, hasErrors } from '../features/forms/validate';
import { LOCKABLE_FIELDS } from '../features/forms/links/linkRules';
import { isLinkCode } from '../features/forms/links/linkCode';
import { withEntity } from '../features/forms/formOptions';

/**
 * A shared checklist, opened from its link by somebody with no sign-in.
 *
 * What may be changed, and what is believed when it comes back, is the
 * server's decision (`server/checklistLinkApi.js`). This page only draws what
 * the server said: IT's fixed values as text, the rest as inputs. Once signed,
 * the same link shows the signed copy, ready to print or save as PDF.
 */

const codeFromPath = () => {
  const match = /^\/c\/([^/?#]+)/.exec(window.location.pathname);
  return match && isLinkCode(match[1]) ? match[1] : null;
};

const GONE = { state: 'gone' };

async function request(code, options) {
  const response = await fetch(`/api/c/${code}`, {
    ...options,
    headers: { Accept: 'application/json', ...(options?.body ? { 'Content-Type': 'application/json' } : null) },
    cache: 'no-store',
    referrerPolicy: 'no-referrer',
  });
  let body = {};
  try {
    body = await response.json();
  } catch {
    // An answer that is not JSON is the host failing, not the form.
  }
  return { status: response.status, body };
}

function Frame({ children }) {
  return (
    <div className="pc-page">
      <header className="pc-bar">
        <Logo size={32} />
        <div className="pc-bar-text">
          <span className="pc-bar-org">PMW Group · IT</span>
          <span className="pc-bar-title">Asset checklist</span>
        </div>
      </header>
      <main className="pc-main">{children}</main>
      <footer className="pc-foot">
        This form was shared with you by PMW IT. If it was not meant for you, please close it.
      </footer>
    </div>
  );
}

function Message({ icon: Icon = AlertTriangle, title, children, action }) {
  return (
    <Card className="ff-done pc-message">
      <Icon size={28} />
      <h1>{title}</h1>
      <p>{children}</p>
      {action}
    </Card>
  );
}

function SignedCopy({ copy, justSigned }) {
  return (
    <>
      <div className="pc-signed-bar">
        <p className="pc-signed-note">
          {justSigned ? <Check size={16} /> : <Lock size={16} />}
          <span>
            {justSigned ? 'Thank you — signed and sent to IT. ' : 'This checklist has been signed and is locked. '}
            Keep a copy: <strong>Print</strong>, then choose <strong>Save as PDF</strong>.
          </span>
        </p>
        <Button icon={Printer} onClick={() => window.print()}>Print / Save as PDF</Button>
      </div>
      <ChecklistRecord
        formMode={copy.formMode}
        values={copy.values}
        signature={copy.signature}
        signedOn={copy.signedOn}
        options={copy.options}
      />
    </>
  );
}

function OpenForm({ code, data, onSigned, onGone }) {
  const [values, setValues] = useState(() => ({
    ...data.values,
    items: data.values.items?.length ? data.values.items : [newItemRow()],
    signature: null,
  }));
  const [errors, setErrors] = useState({});
  const [sending, setSending] = useState(false);
  const [failure, setFailure] = useState('');

  const locked = LOCKABLE_FIELDS.filter((field) => !data.editable.includes(field));
  const mode = FORM_MODES.find((entry) => entry.value === data.formMode);

  const update = (field) => (value) => {
    if (locked.includes(field)) return;
    const next = field === 'entity'
      ? withEntity(values, value, data.options)
      : { ...values, [field]: value };
    setValues(next);
    if (hasErrors(errors)) setErrors(validateChecklist(next));
  };

  const submit = async () => {
    const found = validateChecklist(values);
    setErrors(found);
    if (hasErrors(found)) return;

    setSending(true);
    setFailure('');
    try {
      const { status, body } = await request(code, {
        method: 'POST',
        body: JSON.stringify({ values }),
      });
      if (status === 200 && body.state === 'signed') {
        onSigned(body);
      } else if (status === 422 && body.errors) {
        setErrors(body.errors);
      } else if (status === 404) {
        onGone();
      } else if (status === 409 && body.state === 'signed') {
        window.location.reload();
      } else {
        setFailure(body.error || 'Your form could not be sent just now. Your answers are still here — please try again.');
      }
    } catch {
      // Never cleared: retyping a checklist because the signal dropped is the
      // worst thing a form can do to somebody.
      setFailure('You seem to be offline. Your answers are still here — try again once you have a connection.');
    } finally {
      setSending(false);
    }
  };

  return (
    <>
      <div className="pc-intro">
        <h1>IT Asset Tracking Form</h1>
        <p>
          Check the details below, fill in anything still empty, then sign.
          Fields with a <Lock size={12} aria-label="lock" /> were filled in by IT.
        </p>
      </div>

      {failure && <ErrorBanner message={failure} onRetry={submit} />}

      <Card className="ff-panel">
        <div className="pc-mode">
          <span className="pc-mode-label">{mode?.label ?? data.formMode}</span>
          {mode?.description && <span className="pc-mode-desc">{mode.description}</span>}
        </div>

        <ChecklistFields
          values={values}
          errors={errors}
          update={update}
          locked={locked}
          options={data.options}
        />

        <ChecklistSignature
          value={values.signature}
          onChange={update('signature')}
          error={errors.signature}
        />

        {hasErrors(errors) && (
          <p className="ff-summary" role="alert">
            <AlertTriangle size={14} />
            Some answers are still needed — they are marked above.
          </p>
        )}

        <div className="ff-wizard-foot">
          <p className="pc-commit">
            Once you submit, this checklist is locked and cannot be changed.
          </p>
          <Button icon={Check} onClick={submit} disabled={sending}>
            {sending ? 'Sending…' : 'Sign and submit'}
          </Button>
        </div>
      </Card>
    </>
  );
}

export default function PublicChecklist() {
  const [code] = useState(codeFromPath);
  const [data, setData] = useState(() => (code ? null : GONE));
  const [justSigned, setJustSigned] = useState(false);
  const [attempt, setAttempt] = useState(0);

  useEffect(() => {
    document.title = 'PMW IT — Asset checklist';
  }, []);

  useEffect(() => {
    if (!code) return undefined;
    let live = true;
    request(code).then(
      ({ status, body }) => {
        if (!live) return;
        if (status === 200) setData(body);
        else if (status === 404) setData(GONE);
        else setData({ state: 'error', error: body.error });
      },
      () => {
        if (live) setData({ state: 'error' });
      },
    );
    return () => { live = false; };
  }, [code, attempt]);

  const retry = () => {
    setData(null);
    setAttempt((count) => count + 1);
  };

  let body;
  if (!data) {
    body = (
      <Card className="ff-progress"><span className="spinner" /> Opening your checklist…</Card>
    );
  } else if (data.state === 'gone') {
    body = (
      <Message title="This link is no longer available">
        It may have expired or been cancelled. Please ask IT to send you a new one.
      </Message>
    );
  } else if (data.state === 'error') {
    body = (
      <Message
        title="This form cannot be opened right now"
        action={<Button onClick={retry}>Try again</Button>}
      >
        {data.error || 'Please check your connection and try again in a moment.'}
      </Message>
    );
  } else if (data.state === 'signed') {
    body = <SignedCopy copy={data} justSigned={justSigned} />;
  } else {
    body = (
      <OpenForm
        code={code}
        data={data}
        onSigned={(copy) => {
          setJustSigned(true);
          setData(copy);
          window.scrollTo({ top: 0 });
        }}
        onGone={() => setData(GONE)}
      />
    );
  }

  return <Frame>{body}</Frame>;
}
