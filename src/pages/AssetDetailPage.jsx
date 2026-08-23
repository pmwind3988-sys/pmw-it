import { useMemo, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ArrowLeft, Save, Trash2, Tag, Barcode, Check,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { updateAsset, deleteAsset, EDITABLE_FIELDS } from '../features/assets/sharepoint/updateAsset';
import { CATEGORIES, CONDITIONS, STATUSES, TRACKED, BULK } from '../features/assets/assetKinds';
import { formatMYT } from '../features/datastudio/time/malaysiaTime';

/**
 * One item, in full, and editable.
 *
 * Everything is editable here, unlike the device register — there is no scan
 * file to disagree with. A barcode said what it said, and a person holding the
 * thing knows better.
 */

const FIELDS = [
  { key: 'category', label: 'Category', options: CATEGORIES },
  { key: 'trackingMode', label: 'Counted as', options: [TRACKED, BULK] },
  { key: 'manufacturer', label: 'Make' },
  { key: 'model', label: 'Model' },
  { key: 'serialNumber', label: 'Serial number' },
  { key: 'partNumber', label: 'Part number' },
  { key: 'macAddress', label: 'MAC address' },
  { key: 'assetTag', label: 'Asset label' },
  { key: 'quantity', label: 'Quantity', type: 'number' },
  { key: 'condition', label: 'Condition', options: CONDITIONS },
  { key: 'status', label: 'Status', options: STATUSES },
  { key: 'location', label: 'Where it is' },
  { key: 'supplier', label: 'Purchased from' },
  { key: 'poNumber', label: 'PO number' },
  { key: 'specSummary', label: 'Specification', multiline: true },
  { key: 'remarks', label: 'Remarks', multiline: true },
];

export default function AssetDetailPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, loading, reload } = useAssets();

  const asset = useMemo(
    () => assets.find((row) => String(row.id) === String(id)),
    [assets, id],
  );

  const [edits, setEdits] = useState({});
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [saved, setSaved] = useState(false);

  // Cleared when the row itself changes, so opening a second item does not
  // arrive carrying the first one's unsaved edits.
  //
  // Adjusted during render rather than in an effect: an effect would paint one
  // frame of the new item wearing the old item's edits, and `npm run lint`
  // fails a setState called straight out of an effect body.
  const [shownId, setShownId] = useState(id);
  if (shownId !== id) {
    setShownId(id);
    setEdits({});
    setSaved(false);
  }

  const valueOf = (key) => (key in edits ? edits[key] : (asset?.[key] ?? ''));
  const dirty = EDITABLE_FIELDS.some(
    (key) => key in edits && String(edits[key]) !== String(asset?.[key] ?? ''),
  );

  const save = async () => {
    setSaving(true);
    setError('');
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      await updateAsset({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        existing: asset,
        edits,
        changedBy: account?.username ?? account?.name ?? '',
      });
      setEdits({});
      setSaved(true);
      reload();
    } catch (failure) {
      setError(failure.message || 'The change could not be saved');
    } finally {
      setSaving(false);
    }
  };

  const remove = async () => {
    // Deleting a row is not undoable and the change log is the only trace, so
    // it asks first.
    if (!window.confirm(`Remove "${asset.title}" from the register?`)) return;

    setSaving(true);
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      await deleteAsset({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        asset,
        changedBy: account?.username ?? account?.name ?? '',
      });
      reload();
      navigate('/assets');
    } catch (failure) {
      setError(failure.message || 'The item could not be removed');
      setSaving(false);
    }
  };

  if (loading) return <AppShell title="Item"><div className="spinner" /></AppShell>;

  if (!asset) {
    return (
      <AppShell title="Item">
        <EmptyState>That item is not in the register.</EmptyState>
        <Button variant="secondary" icon={ArrowLeft} onClick={() => navigate('/assets')}>
          Back to the register
        </Button>
      </AppShell>
    );
  }

  return (
    <AppShell
      title={asset.title || 'Item'}
      subtitle={[asset.category, asset.assetTag && `label ${asset.assetTag}`]
        .filter(Boolean).join(' · ')}
      actions={(
        <>
          <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/assets')}>Back</Button>
          <Button icon={Save} onClick={save} disabled={!dirty || saving}>
            {saving ? 'Saving…' : 'Save changes'}
          </Button>
        </>
      )}
    >
      {error && <ErrorBanner message={error} onRetry={save} />}

      {saved && !dirty && (
        <Card className="as-notice as-notice-ok">
          <Check size={16} />
          <span>Saved. The change is recorded in the asset change log.</span>
        </Card>
      )}

      <div className="as-detail">
        <Card className="as-panel">
          <h2 className="as-h2">Details</h2>
          <div className="as-form">
            {FIELDS.map((field) => (
              <label className="as-field" key={field.key}>
                <span className="as-field-label">
                  {field.label}
                  {asset.manualFields?.includes(field.key) && (
                    <span className="as-guess" title="Set by hand, so a re-scan will not change it">
                      hand-set
                    </span>
                  )}
                </span>

                {field.options ? (
                  <select
                    value={valueOf(field.key)}
                    onChange={(e) => setEdits({ ...edits, [field.key]: e.target.value })}
                  >
                    <option value="">—</option>
                    {field.options.map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                ) : field.multiline ? (
                  <textarea
                    rows={3}
                    value={valueOf(field.key)}
                    onChange={(e) => setEdits({ ...edits, [field.key]: e.target.value })}
                  />
                ) : (
                  <input
                    type={field.type ?? 'text'}
                    value={valueOf(field.key)}
                    onChange={(e) => setEdits({ ...edits, [field.key]: e.target.value })}
                  />
                )}
              </label>
            ))}
          </div>
        </Card>

        <div className="as-detail-side">
          {asset.photoUrl && (
            <Card className="as-panel">
              <h2 className="as-h2">Photo</h2>
              <img src={asset.photoUrl} alt={asset.title} className="as-detail-photo" />
            </Card>
          )}

          <Card className="as-panel">
            <h2 className="as-h2">Where it came from</h2>
            <dl className="as-facts">
              <dt>Delivery</dt>
              <dd>{asset.batchTitle || '—'}</dd>
              <dt>Arrived</dt>
              <dd>
                {asset.arrivedOnMYT
                  || (asset.arrivedOn ? formatMYT(asset.arrivedOn, 'datetime12') : '—')}
              </dd>
              <dt>Added</dt>
              <dd>{asset.addedOnMYT || '—'}</dd>
              <dt>Added by</dt>
              <dd>{asset.addedBy || '—'}</dd>
              {asset.poPhotoUrl && (
                <>
                  <dt>PO scan</dt>
                  <dd><a href={asset.poPhotoUrl} className="as-link">Open</a></dd>
                </>
              )}
            </dl>
          </Card>

          <Card className="as-panel">
            <h2 className="as-h2">Identity</h2>
            <p className="as-mono as-key">
              <Barcode size={13} /> {asset.assetKey || '—'}
            </p>
            {asset.assetTag && (
              <p className="as-mono"><Tag size={13} /> {asset.assetTag}</p>
            )}
            {asset.additionalCodes?.length > 0 && (
              <p className="as-sub">Also read: {asset.additionalCodes.join(', ')}</p>
            )}
          </Card>

          <Button variant="ghost" icon={Trash2} onClick={remove} disabled={saving}>
            Remove from the register
          </Button>
        </div>
      </div>
    </AppShell>
  );
}
