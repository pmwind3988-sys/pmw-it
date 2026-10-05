import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import QRCode from 'qrcode';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner } from '../components/ui/Surfaces';
import { Copy, Check, Link2, ClipboardList, Plus } from '../components/ui/Icons';
import Field from '../components/form/Field';
import { RadioCards } from '../components/form/Choices';
import { SelectInput } from '../components/form/Inputs';
import ChecklistFields from '../components/checklist/ChecklistFields';
import { MODE_ICONS } from '../components/checklist/modeIcons';
import { useSharePointToken } from '../hooks/useRequests';
import { useOrgDirectory } from '../hooks/useOrgDirectory';
import { snapshotOptions, withEntity } from '../features/forms/formOptions';
import { FORM_MODES, emptyChecklist, modeLabel } from '../features/forms/checklistForm';
import { isBlankValue } from '../features/forms/links/linkRules';
import { createLink, draftLink, linkUrl } from '../features/forms/sharepoint/checklistLinks';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const EXPIRY_OPTIONS = [
  { value: '3', label: '3 days' },
  { value: '7', label: '7 days' },
  { value: '14', label: '14 days' },
  { value: '30', label: '30 days' },
];

/**
 * Pre-filling an asset checklist for somebody else to sign.
 *
 * IT fills what it knows and decides, field by field, whether the employee may
 * change it; anything left blank is the employee's to fill regardless. The
 * result is a short link to copy into a chat — nothing is sent from here.
 */
function EditSwitch({ field, values, editable, onToggle }) {
  if (isBlankValue(field, values[field])) {
    return <span className="ff-switch ff-switch-fixed">Employee fills this in</span>;
  }
  // A department belongs to a company: an entity the employee can change
  // brings its department with it (`employeeMayEdit`).
  if (field === 'department'
    && (isBlankValue('entity', values.entity) || editable.includes('entity'))) {
    return <span className="ff-switch ff-switch-fixed">Employee can edit, with Entity</span>;
  }
  const on = editable.includes(field);
  return (
    <label className={`ff-switch${on ? ' ff-switch-on' : ''}`}>
      <input
        type="checkbox"
        className="ff-sr-input"
        checked={on}
        onChange={() => onToggle(field)}
      />
      <span className="ff-switch-track" aria-hidden="true" />
      Employee can edit
    </label>
  );
}

function CopyLink({ url }) {
  const [copied, setCopied] = useState(false);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url);
    } catch {
      // Older browsers, or a page not allowed the clipboard: select the text
      // so a long-press or Ctrl+C still works.
      document.getElementById('share-url')?.select();
      return;
    }
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="cl-copy">
      <input
        id="share-url"
        className="ff-input cl-copy-url"
        value={url}
        readOnly
        onFocus={(event) => event.target.select()}
        aria-label="Link to the checklist"
      />
      <Button icon={copied ? Check : Copy} onClick={copy}>
        {copied ? 'Copied' : 'Copy link'}
      </Button>
    </div>
  );
}

export default function ChecklistSharePage() {
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  // HR's entities and departments. The link keeps a copy of them, because the
  // employee opening it cannot read HR's lists.
  const options = snapshotOptions(useOrgDirectory());

  const [values, setValues] = useState(emptyChecklist);
  const [editable, setEditable] = useState([]);
  const [expiry, setExpiry] = useState('14');
  const [busy, setBusy] = useState(false);
  const [failure, setFailure] = useState('');
  const [created, setCreated] = useState(null);
  const [qr, setQr] = useState('');

  useEffect(() => {
    document.title = 'PMW IT — Share a checklist';
  }, []);

  const url = created ? linkUrl(window.location.origin, created.code) : '';

  useEffect(() => {
    if (!url) return undefined;
    let live = true;
    QRCode.toDataURL(url, { margin: 1, width: 220 }).then((data) => {
      if (live) setQr(data);
    }).catch(() => {});
    return () => { live = false; };
  }, [url]);

  const update = (field) => (value) => setValues((current) => (field === 'entity'
    ? withEntity(current, value, options)
    : { ...current, [field]: value }));
  const toggle = (field) => setEditable((current) => (current.includes(field)
    ? current.filter((entry) => entry !== field)
    : [...current, field]));

  const create = async () => {
    if (!values.formMode) {
      setFailure('Pick what this checklist is for first.');
      return;
    }
    setBusy(true);
    setFailure('');
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      const link = await createLink({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        link: draftLink({
          formMode: values.formMode,
          values,
          editable,
          options,
          expiresInDays: Number(expiry),
        }),
        createdByName: account?.name ?? '',
        createdByEmail: account?.username ?? '',
      });
      setCreated(link);
    } catch (thrown) {
      // What IT typed stays on screen, so a retry is one press.
      setFailure(thrown.message || 'The link could not be created');
    } finally {
      setBusy(false);
    }
  };

  const startAgain = () => {
    setValues(emptyChecklist());
    setEditable([]);
    setCreated(null);
    setQr('');
  };

  const listButton = (
    <Button variant="ghost" icon={ClipboardList} onClick={() => navigate('/asset-checklist/links')}>
      Shared checklists
    </Button>
  );

  if (created) {
    return (
      <AppShell title="Checklist link ready" subtitle={`${modeLabel(created.formMode)} — ${created.preset.employeeName || 'no name set'}`} actions={listButton}>
        <Card className="ff-panel cl-ready">
          <div className="cl-ready-head">
            <Link2 size={22} />
            <div>
              <h2 className="ff-panel-title">Copy this link and send it to the employee</h2>
              <p className="ff-help">
                Anyone with the link can open it, so send it only to them. It works until{' '}
                {new Date(created.expiresOn).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}
                {' '}or until they sign. After signing it shows their signed copy.
              </p>
            </div>
          </div>

          <CopyLink url={url} />

          {qr && (
            <figure className="cl-qr">
              <img src={qr} alt="QR code for the checklist link" width="220" height="220" />
              <figcaption>Or let them scan it from this screen.</figcaption>
            </figure>
          )}

          <div className="ff-actions">
            <Button variant="ghost" onClick={() => window.open(url, '_blank', 'noopener')}>Open it</Button>
            <Button icon={Plus} onClick={startAgain}>Share another</Button>
          </div>
        </Card>
      </AppShell>
    );
  }

  return (
    <AppShell
      title="Share a checklist"
      subtitle="Fill in what you know, then copy a link for the employee to complete and sign"
      actions={listButton}
    >
      {failure && <ErrorBanner message={failure} onRetry={values.formMode ? create : undefined} />}

      <Card className="ff-panel">
        <Field label="Form Type" required help="The employee cannot change this.">
          <RadioCards
            name="formMode"
            value={values.formMode}
            onChange={update('formMode')}
            options={FORM_MODES}
            icons={MODE_ICONS}
          />
        </Field>

        {values.formMode && (
          <>
            <p className="ff-help cl-hint">
              Everything is optional. A field you fill is fixed unless you switch on
              {' '}<strong>Employee can edit</strong>; a field you leave blank is theirs to fill.
            </p>

            <ChecklistFields
              values={values}
              update={update}
              requireDetails={false}
              options={options}
              adornment={(field) => (
                <EditSwitch field={field} values={values} editable={editable} onToggle={toggle} />
              )}
            />

            <div className="ff-grid">
              <Field label="Link works for" htmlFor="expiry" help="After this it stops opening, unless it was already signed.">
                <SelectInput
                  id="expiry"
                  value={expiry}
                  onChange={(value) => setExpiry(value || '14')}
                  options={EXPIRY_OPTIONS}
                  placeholder="14 days"
                />
              </Field>
            </div>

            <div className="ff-wizard-foot">
              <Button icon={Link2} onClick={create} disabled={busy}>
                {busy ? 'Creating link…' : 'Create link'}
              </Button>
            </div>
          </>
        )}
      </Card>
    </AppShell>
  );
}
