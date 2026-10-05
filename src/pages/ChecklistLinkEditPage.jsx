import { useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner } from '../components/ui/Surfaces';
import { AlertTriangle, ArrowLeft, Save } from '../components/ui/Icons';
import ChecklistFields from '../components/checklist/ChecklistFields';
import { useSharePointToken } from '../hooks/useRequests';
import { modeLabel, newItemRow } from '../features/forms/checklistForm';
import { LINK_STATUS } from '../features/forms/links/linkSchema';
import { withEntity } from '../features/forms/formOptions';
import { getLink, editSignedLink } from '../features/forms/sharepoint/checklistLinks';
import { formatMYT } from '../utils/malaysiaTime';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * IT correcting a checklist the employee has already signed.
 *
 * The signature cannot be changed here and the form type cannot either —
 * those are what the employee did. Everything else can, and saving says so
 * on the record ("Edited by … after signing"); the values as first signed
 * are kept on the link. What an edit writes is `planEdit`'s decision.
 */
export default function ChecklistLinkEditPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();

  const [link, setLink] = useState(null);
  const [values, setValues] = useState(null);
  const [errors, setErrors] = useState({});
  const [failure, setFailure] = useState('');
  const [saving, setSaving] = useState(false);
  const [loadFailed, setLoadFailed] = useState('');

  useEffect(() => {
    document.title = 'PMW IT — Edit signed checklist';
  }, []);

  useEffect(() => {
    let live = true;
    (async () => {
      try {
        const tokenRes = await getToken();
        const found = await getLink({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, id });
        if (!live) return;
        if (!found) {
          setLoadFailed('This shared checklist no longer exists.');
        } else if (found.status !== LINK_STATUS.SIGNED || !found.submitted) {
          setLoadFailed('Only a signed checklist can be edited. This one has not been signed.');
        } else {
          setLink(found);
          setValues({
            ...found.submitted,
            formMode: found.formMode,
            items: found.submitted.items?.length ? found.submitted.items : [newItemRow()],
          });
        }
      } catch (thrown) {
        if (live) setLoadFailed(thrown.message || 'The shared checklist could not be loaded');
      }
    })();
    return () => { live = false; };
  }, [getToken, id]);

  const update = (field) => (value) => setValues((current) => (field === 'entity'
    ? withEntity(current, value, link?.options ?? null)
    : { ...current, [field]: value }));

  const save = async () => {
    setSaving(true);
    setFailure('');
    try {
      const tokenRes = await getToken();
      const result = await editSignedLink({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        link,
        values,
        by: instance.getActiveAccount()?.name || instance.getActiveAccount()?.username || 'IT',
      });
      if (Object.keys(result.errors).length) {
        setErrors(result.errors);
        return;
      }
      navigate('/asset-checklist/links');
    } catch (thrown) {
      // The corrections stay on screen, so a retry is one press.
      setFailure(thrown.message || 'The changes could not be saved');
    } finally {
      setSaving(false);
    }
  };

  const back = (
    <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/asset-checklist/links')}>
      Shared checklists
    </Button>
  );

  if (loadFailed) {
    return (
      <AppShell title="Edit signed checklist" actions={back}>
        <ErrorBanner message={loadFailed} />
      </AppShell>
    );
  }

  if (!link) {
    return (
      <AppShell title="Edit signed checklist" actions={back}>
        <Card className="ff-progress"><span className="spinner" /> Loading…</Card>
      </AppShell>
    );
  }

  const signedAt = Date.parse(link.signedOn);

  return (
    <AppShell
      title="Edit signed checklist"
      subtitle={`${modeLabel(link.formMode)} — ${link.employeeName || 'no name'}`}
      actions={back}
    >
      {failure && <ErrorBanner message={failure} onRetry={save} />}

      <Card className="ff-panel">
        <p className="cl-warning" role="note">
          <AlertTriangle size={16} />
          <span>
            {link.employeeName || 'The employee'} signed this
            {Number.isFinite(signedAt) ? ` on ${formatMYT(signedAt, 'datetime12')}` : ''}.
            {' '}Saving changes marks the record and the signed copy <strong>“Edited after signing”</strong>
            {' '}with your name. The values as they were first signed are kept. To have them
            sign the corrected version, <strong>Reopen</strong> the link instead.
          </span>
        </p>

        <ChecklistFields
          values={values}
          errors={errors}
          update={update}
          options={link.options ?? null}
        />

        <div className="ff-wizard-foot">
          <Button variant="ghost" onClick={() => navigate('/asset-checklist/links')} disabled={saving}>
            Discard changes
          </Button>
          <Button icon={Save} onClick={save} disabled={saving}>
            {saving ? 'Saving…' : 'Save changes'}
          </Button>
        </div>
      </Card>
    </AppShell>
  );
}
