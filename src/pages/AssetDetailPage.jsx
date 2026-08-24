import { useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ArrowLeft, Save, Trash2, Tag, Barcode, Check, Users, Clock,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { updateAsset, deleteAsset, EDITABLE_FIELDS } from '../features/assets/sharepoint/updateAsset';
import {
  CATEGORIES, CONDITIONS, STATUSES, TRACKED, BULK, isTracked,
} from '../features/assets/assetKinds';
import { formatMYT } from '../features/datastudio/time/malaysiaTime';
import { useHandovers } from '../features/assets/useHandovers';
import {
  holdersOf, outstanding, isOpen, isOverdue, available, owned,
} from '../features/assets/handover/availability';
import {
  unitsOf, serialiseUnits, filledCount, PER_UNIT_ONLY,
} from '../features/assets/units';
import UnitPager from '../features/assets/ui/UnitPager';
import AssetPhoto from '../features/assets/ui/AssetPhoto';
import { absoluteFileUrl } from '../features/assets/sharepoint/fileUrl';

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
  const { handovers } = useHandovers();

  const asset = useMemo(
    () => assets.find((row) => String(row.id) === String(id)),
    [assets, id],
  );

  // Split rather than one list: who has it NOW is the answer somebody opened
  // this page for, and its history is the answer they might scroll to.
  const holders = useMemo(
    () => (asset ? holdersOf(handovers, asset.assetKey) : []),
    [handovers, asset],
  );
  const history = useMemo(
    () => (asset
      ? handovers
        .filter((row) => row.assetKey === asset.assetKey && !isOpen(row))
        .sort((a, b) => (b.issuedOn ?? 0) - (a.issuedOn ?? 0))
      : []),
    [handovers, asset],
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

  // The individual things inside a bulk line. Built from the quantity being
  // SHOWN rather than the one that was saved, so typing 3 into the quantity box
  // grows the pager to three cards before anybody presses Save.
  const units = useMemo(
    () => unitsOf({ ...asset, quantity: Number(valueOf('quantity')) || asset?.quantity },
      'units' in edits ? edits.units : asset?.units),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [asset, edits.units, edits.quantity],
  );
  // Shown for every bulk line, however few are on it — a line of one still
  // holds its serial on the item rather than on the row, and hiding the pager
  // at a quantity of 1 would leave that serial nowhere to go.
  const perUnit = Boolean(asset) && !isTracked(asset.trackingMode);

  // A serial, a part number, a MAC, a label, a condition and a status each
  // describe one physical thing, so a bulk line does not offer them on the
  // row. They are on the items, below.
  const shownFields = perUnit
    ? FIELDS.filter((field) => !PER_UNIT_ONLY.includes(field.key))
    : FIELDS;
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
        {/* Two columns on a desktop: the record itself down the left, and the
            things around it — photo, who has it, where it came from — down the
            right. The unit pager belongs with the record, not beside it. */}
        <div className="as-detail-main">
          <Card className="as-panel">
            <h2 className="as-h2">Details</h2>
            <div className="as-form">
              {shownFields.map((field) => (
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

          {/* A bulk line is a count of identical things, which is the right
              answer to "what did we buy" and the wrong one to "which of the two
              tabs has the cracked screen". This is where each physical item gets
              its own serial, label and condition. */}
          {perUnit && (
            <Card className="as-panel as-units-panel">
              <h2 className="as-h2">
                Each one, individually
                <span className="as-sub">
                  {filledCount(units)} of {units.length} recorded
                </span>
              </h2>
              <UnitPager
                units={units}
                onChange={(next) => setEdits({ ...edits, units: serialiseUnits(next) })}
              />
            </Card>
          )}
        </div>

        <div className="as-detail-side">
          <Card className="as-panel">
            <h2 className="as-h2">Photo</h2>
            <AssetPhoto
              siteUrl={SHAREPOINT_SITE_URL}
              stored={asset.photoUrl}
              alt={asset.title}
            />
          </Card>

          <Card className="as-panel">
            <h2 className="as-h2">Who has it</h2>
            {holders.length === 0 ? (
              <p className="as-sub">
                Nothing is out — {available(asset)} of {owned(asset)} on the shelf.
              </p>
            ) : (
              <ul className="as-holders">
                {holders.map((row) => (
                  <li key={row.id} className={isOverdue(row) ? 'as-row-overdue' : undefined}>
                    <Link
                      to={`/assets/people/${encodeURIComponent(row.personEmail)}`}
                      className="as-link"
                    >
                      <Users size={13} /> {row.personName || row.personEmail}
                    </Link>
                    <span className="as-sub">
                      {outstanding(row)} · {row.kind}
                      {row.dueOnMYT ? ` · due ${row.dueOnMYT}` : ''}
                    </span>
                    {isOverdue(row) && (
                      <span className="as-overdue"><Clock size={12} /> overdue</span>
                    )}
                  </li>
                ))}
              </ul>
            )}
            <Button
              variant="secondary"
              size="sm"
              icon={Users}
              onClick={() => navigate('/assets/handover')}
              disabled={available(asset) < 1}
            >
              Hand this out
            </Button>
          </Card>

          {history.length > 0 && (
            <Card className="as-panel">
              <h2 className="as-h2">Handover history</h2>
              <ul className="as-holders">
                {history.map((row) => (
                  <li key={row.id}>
                    <span>{row.personName || row.personEmail}</span>
                    <span className="as-sub">
                      {row.quantity} · {row.issuedOnMYT || ''}
                      {row.returnedOnMYT ? ` — back ${row.returnedOnMYT}` : ''}
                      {row.returnCondition ? ` (${row.returnCondition})` : ''}
                    </span>
                  </li>
                ))}
              </ul>
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
                  <dd>
                    {/* Absolute, for the same reason the photograph is: the
                        stored path belongs to SharePoint, and a relative link
                        asks the portal for a file it has never had. */}
                    <a
                      href={absoluteFileUrl(SHAREPOINT_SITE_URL, asset.poPhotoUrl)}
                      target="_blank"
                      rel="noreferrer"
                      className="as-link"
                    >
                      Open
                    </a>
                  </dd>
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
