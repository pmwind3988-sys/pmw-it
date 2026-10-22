import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { Link, useNavigate, useParams, useSearchParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { gradeCssVars } from '../features/devices/standards/gradeColors';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { Pencil, RefreshCw, Trash2 } from '../components/ui/Icons';
import { useConfirm } from '../components/ui/useConfirm';
import { useDevices } from '../features/devices/useDevices';
import { groupsFor, RAW_REPORT_KEY } from '../features/devices/fieldGroups';
import { formatScalar } from '../features/devices/formatValue';
import ValueCell from '../features/devices/ui/ValueCell';
import { toneForField, toneForEntry, hasEntryTones } from '../features/devices/fieldTone';
import { formatMYT, toEpochMs } from '../utils/malaysiaTime';
import { relativeDays } from '../utils/relativeDays';
import { useSharePointToken } from '../hooks/useRequests';
import { useDeviceHistory } from '../features/devices/useDeviceHistory';
import LifecycleActions from '../features/devices/ui/LifecycleActions';
import OwnerHistory from '../features/devices/ui/OwnerHistory';
import SpecHistory from '../features/devices/ui/SpecHistory';
import { performLifecycle, retryStints, correctOpenStint } from '../features/devices/sharepoint/writeLifecycle';
import { statusOf } from '../features/devices/lifecycle/status';
import { locationsIn } from '../features/devices/map/locations';
import { mapHref } from '../features/devices/map/mapLinks';
import PartTable from '../features/devices/ui/PartTable';
import PlaceFill from '../features/devices/ui/PlaceFill';
import { updateDevice, deleteDevice, renameDevice } from '../features/devices/sharepoint/updateDevice';
import RenameDialog from '../features/devices/ui/RenameDialog';
import { useDeviceAdmin } from '../features/devices/useDeviceAdmin';
import { matchingGroups } from '../features/devices/fieldSearch';

/** Remembered per browser: somebody who turns the colouring off is not asked
 *  to turn it off again on the next machine they open. */
const TONE_KEY = 'deviceValueTones';

const TABS = [
  ['details', 'Details'],
  ['history', 'History'],
  ['report', 'Scan report'],
];

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const readTonePreference = () => {
  try {
    return localStorage.getItem(TONE_KEY) !== 'off';
  } catch {
    return true;
  }
};

/** A scan's time as an instant, or NaN when the register holds no usable value. */
const instantOf = (value) => {
  if (value === null || value === undefined || value === '') return NaN;
  try {
    const ms = typeof value === 'number' ? value : toEpochMs(value);
    return Number.isFinite(ms) ? ms : NaN;
  } catch {
    return NaN;
  }
};

/** A field's value as the search reads it: what the page would show. */
const searchText = (device, field) => {
  try {
    return formatScalar(device[field.key], field.kind);
  } catch {
    return String(device[field.key] ?? '');
  }
};

/** The serial with a Copy button that says Copied for two seconds. */
function SerialValue({ serial }) {
  const [copied, setCopied] = useState(false);
  const timer = useRef(null);
  const canCopy = typeof navigator !== 'undefined' && Boolean(navigator.clipboard?.writeText);

  useEffect(() => () => clearTimeout(timer.current), []);

  const copy = async () => {
    if (!canCopy) return;
    try {
      await navigator.clipboard.writeText(String(serial));
      setCopied(true);
      clearTimeout(timer.current);
      timer.current = setTimeout(() => setCopied(false), 2000);
    } catch {
      // The serial is still on screen to copy by hand.
    }
  };

  return (
    <div className="mh-def">
      <span className="mh-def-k">Serial</span>
      <span className="mh-def-v mh-mono">{serial}</span>
      <button type="button" className="mh-copy" onClick={copy} disabled={!canCopy}>
        {copied ? 'Copied' : 'Copy'}
      </button>
    </div>
  );
}

/** One machine: its health, its owner and place, and everything the scan read. */
export default function DeviceDetailPage() {
  const { id } = useParams();
  const { devices, loading, error, reload, standards } = useDevices();
  const [params, setParams] = useSearchParams();
  const tabBase = useId();
  const tabRefs = useRef({});
  const [showEmpty, setShowEmpty] = useState(false);
  const [query, setQuery] = useState('');
  const [showTones, setShowTones] = useState(readTonePreference);
  const getToken = useSharePointToken();
  const [acting, setActing] = useState(false);
  const [actionError, setActionError] = useState('');
  const [actionNote, setActionNote] = useState('');
  const [pending, setPending] = useState([]);
  const [placing, setPlacing] = useState(false);
  const [removing, setRemoving] = useState(false);
  const [renameOpen, setRenameOpen] = useState(false);
  const [renaming, setRenaming] = useState(false);
  const [renameError, setRenameError] = useState('');
  const isAdmin = useDeviceAdmin();
  const { ask, dialog } = useConfirm();
  const navigate = useNavigate();

  const requestedTab = params.get('tab');
  const tab = TABS.some(([key]) => key === requestedTab) ? requestedTab : 'details';

  useEffect(() => {
    try {
      localStorage.setItem(TONE_KEY, showTones ? 'on' : 'off');
    } catch {
      // A browser with storage blocked still gets the colours, just not the memory.
    }
  }, [showTones]);

  const device = useMemo(
    () => devices.find((row) => String(row.id) === String(id)),
    [devices, id],
  );

  const history = useDeviceHistory(device);

  const groups = useMemo(
    () => (device ? groupsFor(device, { includeEmpty: showEmpty }) : []),
    [device, showEmpty],
  );
  const shownGroups = useMemo(
    () => matchingGroups(groups, query, (field) => searchText(device ?? {}, field)),
    [groups, query, device],
  );

  const manual = new Set(device?.manualFields ?? []);

  const owners = useMemo(() => [...new Set(devices.map((d) => d.owner).filter(Boolean))].sort(), [devices]);
  const departments = useMemo(() => [...new Set(devices.map((d) => d.department).filter(Boolean))].sort(), [devices]);
  const locations = useMemo(() => locationsIn(devices), [devices]);

  const chooseTab = (key) => {
    setParams((prev) => {
      const next = new URLSearchParams(prev);
      next.set('tab', key);
      return next;
    }, { replace: true });
  };

  const onTabKey = (event) => {
    const at = TABS.findIndex(([key]) => key === tab);
    let to = null;
    if (event.key === 'ArrowRight') to = (at + 1) % TABS.length;
    else if (event.key === 'ArrowLeft') to = (at - 1 + TABS.length) % TABS.length;
    else if (event.key === 'Home') to = 0;
    else if (event.key === 'End') to = TABS.length - 1;
    if (to === null) return;
    event.preventDefault();
    const [key] = TABS[to];
    chooseTab(key);
    tabRefs.current[key]?.focus();
  };

  const act = async (action, input) => {
    setActing(true);
    setActionError('');
    setActionNote('');
    try {
      const tokenRes = await getToken();
      const outcome = await performLifecycle({
        siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, deviceId: device.id, action, input,
        expectedStatus: statusOf(device), recordedBy: tokenRes.account?.username ?? '',
      });
      setPending(outcome.pendingStints);
      if (outcome.pendingStints.length) setActionError('The machine moved, but its owner history could not be written.');
      else if (outcome.logFailed) setActionNote('Saved, but the change history could not be recorded.');
    } catch (failure) {
      setActionError(failure.message);
    } finally {
      setActing(false);
      reload();
      history.reload();
    }
  };

  // An owner, location or department set or corrected by hand: a manual edit
  // so the next import leaves it alone, carried onto the current owner-history
  // entry so the history does not keep the old spelling. Not a move.
  const fillPlace = async (edits) => {
    setPlacing(true);
    setActionError('');
    setActionNote('');
    try {
      const tokenRes = await getToken();
      await updateDevice({
        siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, existing: device, edits,
        changedBy: tokenRes.account?.username ?? '',
      });
      try {
        await correctOpenStint({
          siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, deviceId: device.id, fields: edits,
        });
      } catch {
        // The machine is corrected either way; only the history line lags.
        setActionNote('Saved, but the owner history still shows the old value.');
      }
      return true;
    } catch (failure) {
      setActionError(failure.message);
      return false;
    } finally {
      setPlacing(false);
      reload();
      history.reload();
    }
  };

  const retry = async () => {
    setActing(true);
    try {
      const tokenRes = await getToken();
      await retryStints({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, writes: pending });
      setPending([]);
      setActionError('');
    } catch (failure) {
      setActionError(failure.message);
    } finally {
      setActing(false);
      history.reload();
    }
  };

  // Admin only. The change log keeps a "Removed" line; the owner history rows
  // stay in their own list, tied to an id nothing will use again.
  const remove = async () => {
    if (!await ask({
      title: `Remove ${device.computerName} from the register?`,
      body: 'It disappears from the map, the List and every figure. The change log keeps a '
        + 'line saying it was removed. If the machine still exists, Retire keeps its history instead.',
      confirmLabel: 'Remove it',
      cancelLabel: 'Keep it',
    })) return;
    setRemoving(true);
    setActionError('');
    try {
      const tokenRes = await getToken();
      await deleteDevice({
        siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, device,
        changedBy: tokenRes.account?.username ?? '',
      });
      navigate('/devices?view=register');
    } catch (failure) {
      setActionError(failure.message);
      setRemoving(false);
    }
  };

  // Admin only. The old name stays in the History tab as a "Renamed" line.
  const rename = async (name) => {
    setRenaming(true);
    setRenameError('');
    try {
      const tokenRes = await getToken();
      const { logFailed } = await renameDevice({
        siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, existing: device, newName: name,
        devices, changedBy: tokenRes.account?.username ?? '',
      });
      setRenameOpen(false);
      setActionError('');
      setActionNote(logFailed ? 'Renamed, but the old name could not be written to the history.' : '');
      reload();
      history.reload();
    } catch (failure) {
      setRenameError(failure.message);
    } finally {
      setRenaming(false);
    }
  };

  const toggleTones =() => setShowTones((on) => !on);

  const status = device ? statusOf(device) : '';
  const kind = device ? [device.deviceType, device.computerModel].filter(Boolean).join(' · ') : '';
  const scanned = device ? instantOf(device.scannedOn) : NaN;
  const rawReport = device?.[RAW_REPORT_KEY];

  return (
    <AppShell
      title={device?.computerName ?? 'Device'}
      subtitle={device
        ? [device.owner, device.department].filter(Boolean).join(' · ') || 'No owner recorded'
        : 'Looking this machine up in the register'}
      actions={(
        <>
          <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reload} loading={loading}>
            Refresh
          </Button>
          {isAdmin && device && (
            <Button variant="secondary" size="sm" icon={Pencil}
              onClick={() => { setRenameError(''); setRenameOpen(true); }}>
              Rename
            </Button>
          )}
          {isAdmin && device && (
            <Button variant="danger" size="sm" icon={Trash2} onClick={remove} loading={removing}>
              Remove
            </Button>
          )}
        </>
      )}
    >
      <div className="dv-graded mh-page" style={gradeCssVars(standards.standard.colors)}>
        {(standards.note || standards.error) && (
          <p className="dv-standard-note" role="status">{standards.note || standards.error}</p>
        )}

        {error && <ErrorBanner message={error} onRetry={reload} />}

        {!device ? (
          <Card>
            <EmptyState>
              {loading
                ? 'Loading the register…'
                : 'No device with that id is in the register any more.'}
              {!loading && (
                <>
                  {' '}
                  <Link to="/devices">Back to the device list</Link>
                </>
              )}
            </EmptyState>
          </Card>
        ) : (
          <>
            <nav className="dm-crumbs" aria-label="Breadcrumb">
              <Link to={mapHref()}>Devices</Link>
              {device.location && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location })}>{device.location}</Link></>)}
              {device.location && device.department && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location, department: device.department })}>{device.department}</Link></>)}
              <span aria-hidden="true">›</span><span aria-current="page">{device.computerName}</span>
            </nav>

            <Card className="mh-head">
              <div className="mh-head-top">
                <div className="mh-head-id">
                  <div className="mh-head-line">
                    <span className={`dd-status dd-status-${status.replace(' ', '-').toLowerCase()}`}>{status}</span>
                    {kind && <span className="mh-kind">{kind}</span>}
                  </div>
                  <div className="mh-defs">
                    <PlaceFill
                      key={`${device.id}|${device.owner ?? ''}|${device.location ?? ''}|${device.department ?? ''}`}
                      device={device} owners={owners} locations={locations} departments={departments}
                      onSave={fillPlace} busy={placing} />
                    {device.serialNumber && <SerialValue serial={device.serialNumber} />}
                    {Number.isFinite(scanned) && (
                      <div className="mh-def">
                        <span className="mh-def-k">Scanned</span>
                        <span className="mh-def-v">
                          {relativeDays(scanned)}
                          <span className="mh-muted"> · {formatMYT(scanned, 'datetime12')}</span>
                        </span>
                      </div>
                    )}
                  </div>
                </div>
                <LifecycleActions device={device} owners={owners} locations={locations} departments={departments}
                  onAction={act} busy={acting} showTones={showTones} onToggleTones={toggleTones} />
              </div>
              {actionNote && !actionError && <p className="dd-note" role="status">{actionNote}</p>}
              {actionError && <ErrorBanner message={actionError} busy={acting} onRetry={pending.length ? retry : undefined} />}
            </Card>

            <PartTable device={device} />

            <div role="tablist" aria-label="More about this machine" className="mh-tabs">
              {TABS.map(([key, label]) => (
                <button key={key} type="button" role="tab" id={`${tabBase}-${key}`}
                  aria-selected={tab === key} aria-controls={`${tabBase}-panel`}
                  tabIndex={tab === key ? 0 : -1}
                  ref={(el) => { tabRefs.current[key] = el; }}
                  className={`mh-tab${tab === key ? ' mh-tab-on' : ''}`}
                  onClick={() => chooseTab(key)} onKeyDown={onTabKey}>
                  {label}
                </button>
              ))}
            </div>

            <div role="tabpanel" id={`${tabBase}-panel`} aria-labelledby={`${tabBase}-${tab}`} className="mh-panel">
              {tab === 'details' && (
                <>
                  <div className="mh-details-bar">
                    <input type="search" className="mh-find" value={query}
                      onChange={(event) => setQuery(event.target.value)} aria-label="Find a field"
                      placeholder="Find a field — e.g. IP, AnyDesk, Outlook" />
                    <label className="mh-blanks">
                      <input type="checkbox" checked={showEmpty} onChange={(event) => setShowEmpty(event.target.checked)} />
                      Show blanks
                    </label>
                  </div>

                  {shownGroups.length === 0 && (
                    <p className="mh-none">
                      {query.trim() ? `No field matches “${query.trim()}”.` : 'No fields to show.'}
                    </p>
                  )}

                  <div className="dd-groups">
                    {shownGroups.map((group) => (
                      <Card key={group.id} className="dd-group">
                        <h2 className="dd-group-title">
                          {group.title}
                          {group.hint && <span className="dd-group-hint">{group.hint}</span>}
                        </h2>
                        <dl className="dd-fields">
                          {group.fields.map((field) => (
                            <div className="dd-field" key={field.key}>
                              <dt>
                                {field.label}
                                {manual.has(field.key) && (
                                  <span className="dt-manual" title="Set by hand — imports leave this alone">
                                    edited
                                  </span>
                                )}
                              </dt>
                              <dd>
                                <ValueCell
                                  value={device[field.key]}
                                  fieldKey={field.key}
                                  kind={field.kind}
                                  tone={showTones ? toneForField(device, field.key) : null}
                                  entryTone={showTones && hasEntryTones(field.key)
                                    ? (text) => toneForEntry(field.key, text)
                                    : undefined}
                                />
                              </dd>
                            </div>
                          ))}
                        </dl>
                      </Card>
                    ))}
                  </div>
                </>
              )}

              {tab === 'history' && (
                <>
                  <div className="dd-histories">
                    <OwnerHistory device={device} stints={history.stints} loading={history.loading} />
                    <SpecHistory changes={history.changes} loading={history.loading} />
                  </div>
                  {history.error && <ErrorBanner message={history.error} onRetry={history.reload} />}
                </>
              )}

              {tab === 'report' && (
                rawReport ? (
                  <Card className="mh-report">
                    <p className="mh-report-file">
                      {device.sourceFileName ? formatScalar(device.sourceFileName) : 'Scan report'}
                    </p>
                    <pre className="dd-raw-text mh-raw">{rawReport}</pre>
                  </Card>
                ) : (
                  <p className="mh-none">No scan report was stored for this machine.</p>
                )
              )}
            </div>
          </>
        )}
      </div>
      {dialog}
      {renameOpen && device && (
        <RenameDialog device={device} devices={devices} busy={renaming} error={renameError}
          onSave={rename} onClose={() => setRenameOpen(false)} />
      )}
    </AppShell>
  );
}
