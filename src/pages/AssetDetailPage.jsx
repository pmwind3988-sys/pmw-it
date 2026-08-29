import { useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import Collapsible from '../components/ui/Collapsible';
import { useConfirm } from '../components/ui/useConfirm';
import {
  ArrowLeft, Save, Trash2, Tag, Barcode, Check, Users, Clock, AlertTriangle,
  ClipboardList,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { updateAsset, deleteAsset, EDITABLE_FIELDS } from '../features/assets/sharepoint/updateAsset';
import {
  CONDITIONS, STATUSES, TRACKED, BULK, isTracked,
} from '../features/assets/assetKinds';
import { categoriesIn } from '../features/assets/categories';
import CategoryField from '../features/assets/ui/CategoryField';
import { needsDetails, missingDetails } from '../features/assets/detailsPending';
import { formatMYT } from '../utils/malaysiaTime';
import { useHandovers } from '../features/assets/useHandovers';
import {
  holdersOf, groupHolders, isOpen, available, owned,
} from '../features/assets/handover/availability';
import {
  unitsOf, serialiseUnits, filledCount, parseUnits, PER_UNIT_ONLY,
} from '../features/assets/units';
import UnitPager from '../features/assets/ui/UnitPager';
import QuantityInput from '../features/assets/ui/QuantityInput';
import ScanField from '../features/assets/ui/ScanField';
import TextScanSheet from '../features/assets/ui/TextScanSheet';
import { labelFor } from '../features/assets/scan/fieldLabels';
import { applyScannedFields, SCAN_FIELDS } from '../features/assets/scan/textScan';
import AssetPhoto from '../features/assets/ui/AssetPhoto';
import SignatureShot from '../features/assets/ui/SignatureShot';
import { absoluteFileUrl } from '../features/assets/sharepoint/fileUrl';
import { uploadUnitPhotos } from '../features/assets/sharepoint/uploadUnitPhotos';
import { loadPhoto, deletePhoto } from '../features/assets/store/assetDb';

/**
 * One item, in full, and editable.
 *
 * Everything is editable here, unlike the device register — there is no scan
 * file to disagree with. A barcode said what it said, and a person holding the
 * thing knows better.
 */

const FIELDS = [
  // Its options are worked out per render, from what the register is using.
  { key: 'category', label: 'Category', category: true },
  { key: 'trackingMode', label: 'Counted as', options: [TRACKED, BULK] },
  { key: 'manufacturer', label: 'Make' },
  { key: 'model', label: 'Model' },
  { key: 'serialNumber', label: 'Serial number' },
  { key: 'partNumber', label: 'Part number' },
  { key: 'macAddress', label: 'MAC address' },
  { key: 'assetTag', label: 'Asset label' },
  { key: 'quantity', label: 'How many?', type: 'number' },
  { key: 'condition', label: 'Condition', options: CONDITIONS },
  { key: 'status', label: 'Status', options: STATUSES },
  { key: 'location', label: 'Where it is' },
  { key: 'supplier', label: 'Purchased from' },
  { key: 'poNumber', label: 'PO number' },
  { key: 'specSummary', label: 'Specification', multiline: true },
  { key: 'remarks', label: 'Remarks', multiline: true },
];

/** Photos taken here and not yet uploaded, so the phone's copies can be cleared. */
function pendingPhotoIds(stored) {
  return parseUnits(stored).map((unit) => unit.photoId).filter(Boolean);
}

/**
 * The camera button, on the fields a printed label can actually fill.
 * A quantity, a status or somebody's remark is not on the box.
 */
function Scannable({ field, onScan, children }) {
  if (!SCAN_FIELDS.includes(field.key)) return children;
  return <ScanField label={field.label} onScan={onScan}>{children}</ScanField>;
}

export default function AssetDetailPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, loading, reload } = useAssets();
  const { handovers } = useHandovers();
  const { ask, dialog } = useConfirm();
  const categories = useMemo(() => categoriesIn(assets), [assets]);

  const asset = useMemo(
    () => assets.find((row) => String(row.id) === String(id)),
    [assets, id],
  );

  // Split rather than one list: who has it NOW is the answer somebody opened
  // this page for, and its history is the answer they might scroll to.
  /**
   * One line per PERSON, not one per handover row.
   *
   * Somebody who took five cables on Monday and one more on Wednesday is two
   * rows in the handover list and one person on this panel. Six identical
   * lines under a person's name is a list nobody reads; "Aisyah · 6" is the
   * answer the panel was opened for, and the serials of the ones that have
   * them are still named underneath it.
   */
  const holders = useMemo(
    // The units are handed in so that a handover naming item 4 of a bulk line
    // can say WHICH monitor that is. Before combining, the row title carried
    // the serial by accident; a line of ten has one title and ten serials, so
    // the item records are the only place left that knows.
    () => (asset ? groupHolders(holdersOf(handovers, asset.assetKey), unitsOf(asset)) : []),
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
  // Which field's camera button is open, and what a scan read but would
  // not write over a value that was set by hand.
  const [scanning, setScanning] = useState(null);
  const [heldBack, setHeldBack] = useState([]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [saved, setSaved] = useState(false);
  const [photoWarning, setPhotoWarning] = useState('');
  // Whether the save had to add a missing column to the SharePoint list first.
  // Worth saying once: it explains why that one save took a minute.
  const [repaired, setRepaired] = useState(false);

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

  /**
   * A scan on a saved row goes into the unsaved edits, not into
   * SharePoint. Nothing here is written until Save, which is what makes
   * pointing the camera at the wrong box harmless.
   */
  const useScan = (values, guessedFields, additional) => {
    const shown = Object.fromEntries(SCAN_FIELDS.map((key) => [key, valueOf(key)]));
    const result = applyScannedFields(
      { ...shown, guessed: asset?.guessed ?? [], manualFields: asset?.manualFields ?? [] },
      values,
      guessedFields,
      additional,
      // Nothing reaches here now except by being ticked off the list.
      { byHand: true },
    );

    const next = { ...edits };
    for (const key of SCAN_FIELDS) {
      if (result.record[key] !== shown[key]) next[key] = result.record[key];
    }

    setEdits(next);
    setHeldBack(result.heldBack);
    // The sheet stays open: a label carries several values, and closing after
    // the first would mean reopening the camera for each of them.
  };

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
  // What the banner lists. Read off the values being SHOWN, so ticking the
  // last serial into the pager empties the list before anything is saved.
  const missing = useMemo(
    () => (asset ? missingDetails({ ...asset, ...edits }) : []),
    [asset, edits],
  );

  const dirty = EDITABLE_FIELDS.some(
    (key) => key in edits && String(edits[key]) !== String(asset?.[key] ?? ''),
  );

  const save = async (pending = edits) => {
    setSaving(true);
    setError('');
    setPhotoWarning('');
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      const token = tokenRes.accessToken;

      // Photographs first, because a unit record must reach SharePoint holding
      // a library path and not a reference to a blob on this phone — which is
      // meaningless to everyone else who opens the row.
      let next = pending;
      let stranded = 0;
      if ('units' in pending) {
        const photos = await uploadUnitPhotos({
          siteUrl: SHAREPOINT_SITE_URL,
          token,
          stored: pending.units,
          seed: asset.assetKey || asset.title,
          photoFor: loadPhoto,
        });
        next = { ...pending, units: photos.units };
        stranded = photos.failures.length;

        if (stranded) {
          setPhotoWarning(
            `${stranded} photo${stranded === 1 ? '' : 's'} could not be uploaded. `
            + 'Everything typed in was saved, and the photo is still on this phone — '
            + 'press Save changes again to try it once more.',
          );
        }
      }

      const result = await updateAsset({
        siteUrl: SHAREPOINT_SITE_URL,
        token,
        existing: asset,
        edits: next,
        changedBy: account?.username ?? account?.name ?? '',
      });

      // The phone's copies are only worth keeping until they are somewhere
      // everybody can see them.
      const stillWaiting = pendingPhotoIds(next.units);
      for (const id of pendingPhotoIds(pending.units)) {
        if (!stillWaiting.includes(id)) await deletePhoto(id).catch(() => {});
      }

      setRepaired(Boolean(result?.repaired));
      // A photo that would not upload keeps its edit, or "press Save again"
      // would be an instruction to press a button that is now greyed out with
      // nothing behind it. Everything else is saved and is cleared.
      setEdits(stranded ? { units: next.units } : {});
      setSaved(true);
      reload();
    } catch (failure) {
      setError(failure.message || 'The change could not be saved');
    } finally {
      setSaving(false);
    }
  };

  /**
   * The row has been finished off. Saved in one press rather than by setting
   * the flag and waiting for the person to find Save changes -- this button
   * IS the save, and `save` is handed the edit directly because state set here
   * would not be readable until the next render.
   */
  const completeDetails = () => save({ ...edits, detailsPending: false });

  const remove = async () => {
    // Deleting a row is not undoable and the change log is the only trace, so
    // it asks first.
    if (!await ask({
      title: `Remove "${asset.title}" from the register?`,
      body: 'Its photographs, its item records and everything typed on it go with it. '
        + 'The change log keeps a line saying it was removed, and that is the only '
        + 'trace left.',
      confirmLabel: 'Remove it',
      cancelLabel: 'Keep it',
    })) return;

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
      // Everything on this page is an edit waiting to be saved, and the page
      // is long enough that Save changes would otherwise scroll away.
      stickyHead
      title={asset.title || 'Item'}
      subtitle={[asset.category, asset.assetTag && `label ${asset.assetTag}`]
        .filter(Boolean).join(' · ')}
      actions={(
        <>
          <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/assets')}>Back</Button>
          {/* Wrapped, not passed: a button hands its click event to whatever
              it is given, and `save(event)` would treat that event as the
              edits -- saving nothing while clearing everything typed in. */}
          <Button icon={Save} onClick={() => save()} disabled={!dirty || saving}>
            {saving ? 'Saving…' : 'Save changes'}
          </Button>
        </>
      )}
    >
      {error && <ErrorBanner message={error} onRetry={() => save()} />}

      {saved && !dirty && (
        <Card className="as-notice as-notice-ok">
          <Check size={16} />
          <span>
            Saved. The change is recorded in the asset change log.
            {repaired && ' The register was missing a column for the item records, '
              + 'so it was added first — later saves will be quick.'}
          </span>
        </Card>
      )}

      {/* A row from a delivery entered long after it arrived. It says what is
          still to be found rather than only that something is, because "needs
          details" on its own sends somebody back to the shelf to work out
          which detail. */}
      {needsDetails(asset) && (
        <Card className="as-notice as-notice-warn as-pending">
          <ClipboardList size={16} />
          <span>
            <strong>Still to be filled in.</strong>{' '}
            {missing.length
              ? `Missing: ${missing.join(', ')}.`
              : 'Everything on this row is filled in now.'}
          </span>
          <Button
            variant="secondary"
            size="sm"
            disabled={saving}
            onClick={completeDetails}
          >
            Details are complete
          </Button>
        </Card>
      )}

      {photoWarning && (
        <Card className="as-notice as-notice-warn">
          <AlertTriangle size={16} />
          <span>{photoWarning}</span>
        </Card>
      )}

      <div className="as-detail">
        {/* Two columns on a desktop: the record itself down the left, and the
            things around it — photo, who has it, where it came from — down the
            right. The unit pager belongs with the record, not beside it. */}
        <div className="as-detail-main">
          {/* Sixteen boxes is the whole item and most of a phone screen. It
              folds, and the fold is remembered — somebody who came to look at
              the items or at who has it should not have to scroll past the
              form every time. The header says what the row IS while it is
              shut, so folding it away costs nothing at a glance. */}
          <Collapsible
            id="asset-details"
            className="as-panel as-fold-panel"
            title="Details"
            summary={[asset.manufacturer, asset.model, asset.category]
              .filter(Boolean).join(' · ') || 'Nothing filled in yet'}
          >
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

                  {field.category ? (
                    // The one dropdown that can be added to. Everything else
                    // here is a fixed list where a new value would be a typo.
                    <CategoryField
                      value={valueOf(field.key)}
                      options={categories}
                      onChange={(next) => setEdits((current) => ({ ...current, [field.key]: next }))}
                    />
                  ) : field.options ? (
                    <select
                      value={valueOf(field.key)}
                      onChange={(e) => setEdits((current) => ({ ...current, [field.key]: e.target.value }))}
                    >
                      <option value="">—</option>
                      {field.options.map((option) => (
                        <option key={option} value={option}>{option}</option>
                      ))}
                    </select>
                  ) : field.multiline ? (
                    // Remarks are somebody's sentence about the thing, so
                    // there is nothing on a label to read into them.
                    <Scannable field={field} onScan={() => setScanning(field.key)}>
                      <textarea
                        rows={3}
                        value={valueOf(field.key)}
                        onChange={(e) => setEdits((current) => ({ ...current, [field.key]: e.target.value }))}
                      />
                    </Scannable>
                  ) : field.type === 'number' ? (
                    // A count that can be emptied while it is being replaced.
                    <QuantityInput
                      value={valueOf(field.key)}
                      onCommit={(count) => setEdits(
                        (current) => ({ ...current, [field.key]: count }),
                      )}
                    />
                  ) : (
                    <Scannable field={field} onScan={() => setScanning(field.key)}>
                      <input
                        type={field.type ?? 'text'}
                        value={valueOf(field.key)}
                        onChange={(e) => setEdits((current) => ({ ...current, [field.key]: e.target.value }))}
                      />
                    </Scannable>
                  )}
                </label>
              ))}
            </div>

            {/* Read off the label but not written, because the value in
                the box was set by hand. Offered rather than dropped. */}
            {heldBack.length > 0 && (
              <ul className="as-heldback">
                {heldBack.map((entry) => (
                  <li key={entry.field}>
                    <span className="as-heldback-label">
                      {labelFor(entry.field)}
                    </span>
                    <span className="as-heldback-value">{entry.value}</span>
                    <button
                      type="button"
                      className="as-heldback-take"
                      onClick={() => {
                        setEdits((current) => ({ ...current, [entry.field]: entry.value }));
                        setHeldBack(heldBack.filter((held) => held.field !== entry.field));
                      }}
                    >
                      Use this instead
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </Collapsible>

          {scanning && (
            <TextScanSheet
              title={`Scan the label — ${labelFor(scanning)}`}
              onCancel={() => setScanning(null)}
              onUse={useScan}
            />
          )}

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
                onChange={(next) => setEdits((current) => ({ ...current, units: serialiseUnits(next) }))}
                siteUrl={SHAREPOINT_SITE_URL}
                rowPhoto={asset.photoUrl}
                poPhoto={asset.poPhotoUrl}
              />
            </Card>
          )}
        </div>

        <div className="as-detail-side">
          {/* Only the paperwork now. The picture of the whole line used to sit
              above it and was worth very little: on a bulk row it is one
              photograph standing in for twenty things, each of which has its
              own picture on its own item card below, and on a tracked row it
              repeats what the item card already shows. The delivery order is
              the one picture with nothing else showing it — the paper saying
              what was SUPPOSED to arrive, which is exactly what somebody
              counting a delivery wants without signing into SharePoint. */}
          <Collapsible
            id="asset-paperwork"
            className="as-panel as-fold-panel"
            title="Delivery order / PO"
            summary={asset.poPhotoUrl ? 'Scanned' : 'Nothing scanned'}
          >
            <div className="as-shots">
              <AssetPhoto
                siteUrl={SHAREPOINT_SITE_URL}
                stored={asset.poPhotoUrl}
                alt="Delivery order"
                caption="Delivery order / PO"
                empty="No delivery order was scanned for this one."
              />
            </div>
          </Collapsible>

          <Card className="as-panel">
            <h2 className="as-h2">Who has it</h2>
            {holders.length === 0 ? (
              <p className="as-sub">
                Nothing is out — {available(asset)} of {owned(asset)} on the shelf.
              </p>
            ) : (
              <ul className="as-holders">
                {holders.map((person) => (
                  <li key={person.key} className={person.overdue ? 'as-row-overdue' : undefined}>
                    <Link
                      to={`/assets/people/${encodeURIComponent(person.email)}`}
                      className="as-link"
                    >
                      <Users size={13} /> {person.name}
                      {/* The count belongs beside the name: one person with six
                          cables is the fact, and six lines saying "1" was not. */}
                      {person.units > 1 && <span className="as-count">× {person.units}</span>}
                    </Link>
                    <span className="as-sub">
                      {person.serials.length ? `${person.serials.join(', ')} · ` : ''}
                      {person.units} · {person.kinds.join(' & ')}
                      {person.dueOnMYT ? ` · due ${person.dueOnMYT}` : ''}
                      {person.lines > 1 ? ` · ${person.lines} handovers` : ''}
                    </span>
                    {person.overdue && (
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
                    {/* Each time this thing changed hands, signed for by the
                        hand it went to. The history of one laptop is exactly
                        where those signatures are wanted, and it used to show
                        none of them. */}
                    <span className="as-sigpair">
                      <SignatureShot
                        stored={row.issueSignature}
                        when="signed on the way out"
                        by={row.personName || row.personEmail}
                      />
                      <SignatureShot
                        stored={row.returnSignature}
                        when="signed on the way back"
                        by={row.personName || row.personEmail}
                      />
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
              {asset.poNumber && (
                <>
                  <dt>PO number</dt>
                  <dd>{asset.poNumber}</dd>
                </>
              )}
              {asset.poPhotoUrl && (
                <>
                  <dt>PO scan</dt>
                  <dd>
                    {/* The scan itself is up with the photographs now. This
                        stays for anyone who wants the file rather than the
                        picture — absolute, because the stored path belongs to
                        SharePoint and a relative link asks the portal for a
                        file it has never had. */}
                    <a
                      href={absoluteFileUrl(SHAREPOINT_SITE_URL, asset.poPhotoUrl)}
                      target="_blank"
                      rel="noreferrer"
                      className="as-link"
                    >
                      Open in SharePoint
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

      {dialog}
    </AppShell>
  );
}
