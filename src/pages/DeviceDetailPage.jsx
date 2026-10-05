import { useEffect, useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { ArrowLeft, RefreshCw } from '../components/ui/Icons';
import { useDevices } from '../features/devices/useDevices';
import { groupsFor, RAW_REPORT_KEY } from '../features/devices/fieldGroups';
import { formatScalar } from '../features/devices/formatValue';
import ValueCell from '../features/devices/ui/ValueCell';
import { toneForField, toneForEntry, hasEntryTones } from '../features/devices/fieldTone';
import { formatMYT } from '../utils/malaysiaTime';
import { useSharePointToken } from '../hooks/useRequests';
import { useDeviceHistory } from '../features/devices/useDeviceHistory';
import LifecycleActions from '../features/devices/ui/LifecycleActions';
import OwnerHistory from '../features/devices/ui/OwnerHistory';
import SpecHistory from '../features/devices/ui/SpecHistory';
import { performLifecycle, retryStints } from '../features/devices/sharepoint/writeLifecycle';
import { statusOf } from '../features/devices/lifecycle/status';
import { locationsIn } from '../features/devices/map/locations';
import { labelOf } from '../features/devices/deviceFilters';
import { mapHref } from '../features/devices/map/mapLinks';

/**
 * The verdict shares the risk palette rather than a second one: red is "go and
 * look", amber "put it on the list", green "leave it alone".
 */
const FIT_TONE = {
  Critical: 'critical',
  'Needs Attention': 'watch',
  Optimal: 'ok',
};

/** Remembered per browser: somebody who turns the colouring off is not asked
 *  to turn it off again on the next machine they open. */
const TONE_KEY = 'deviceValueTones';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const readTonePreference = () => {
  try {
    return localStorage.getItem(TONE_KEY) !== 'off';
  } catch {
    return true;
  }
};

/** One machine, everything the scan read out of it, grouped by what it is. */
export default function DeviceDetailPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { devices, loading, error, reload } = useDevices();
  const [showEmpty, setShowEmpty] = useState(false);
  const [showRaw, setShowRaw] = useState(false);
  const [showTones, setShowTones] = useState(readTonePreference);
  const getToken = useSharePointToken();
  const [acting, setActing] = useState(false);
  const [actionError, setActionError] = useState('');
  const [pending, setPending] = useState([]);

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

  const manual = new Set(device?.manualFields ?? []);

  const owners = useMemo(() => [...new Set(devices.map((d) => d.owner).filter(Boolean))].sort(), [devices]);
  const departments = useMemo(() => [...new Set(devices.map((d) => d.department).filter(Boolean))].sort(), [devices]);
  const locations = useMemo(() => locationsIn(devices), [devices]);

  const act = async (action, input) => {
    setActing(true);
    setActionError('');
    try {
      const tokenRes = await getToken();
      const outcome = await performLifecycle({
        siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, deviceId: device.id, action, input,
        expectedStatus: statusOf(device), recordedBy: tokenRes.account?.username ?? '',
      });
      setPending(outcome.pendingStints);
      if (outcome.pendingStints.length) setActionError('The machine moved, but its owner history could not be written.');
    } catch (failure) {
      setActionError(failure.message);
    } finally {
      setActing(false);
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

  return (
    <AppShell
      title={device?.computerName ?? 'Device'}
      subtitle={device
        ? [device.owner, device.department].filter(Boolean).join(' · ') || 'No owner recorded'
        : 'Looking this machine up in the register'}
      actions={(
        <>
          <Button variant="secondary" size="sm" icon={ArrowLeft} onClick={() => navigate(-1)}>
            Back
          </Button>
          <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reload} loading={loading}>
            Refresh
          </Button>
        </>
      )}
    >
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
                <Link to="/devices">Back to the map</Link>
              </>
            )}
          </EmptyState>
        </Card>
      ) : (
        <>
          <nav className="dm-crumbs" aria-label="Breadcrumb">
            <Link to={mapHref()}>Map</Link>
            {device.location && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location })}>{device.location}</Link></>)}
            {device.location && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location, department: labelOf(device.department) })}>{labelOf(device.department)}</Link></>)}
            <span aria-hidden="true">›</span><span aria-current="page">{device.computerName}</span>
          </nav>

          <Card className="dd-life">
            <div className="dd-life-head">
              <span className={`dd-status dd-status-${statusOf(device).replace(' ', '-').toLowerCase()}`}>{statusOf(device)}</span>
              <span className="dd-life-who">
                {device.owner
                  ? <>With <strong>{device.owner}</strong>{[device.location, device.department].filter(Boolean).map((v) => ` · ${v}`).join('')}</>
                  : 'Nobody has it'}
                {device.serialNumber && <span className="dd-life-serial">Serial {device.serialNumber}</span>}
              </span>
              <LifecycleActions device={device} owners={owners} locations={locations} departments={departments} onAction={act} busy={acting} />
            </div>
            {actionError && <ErrorBanner message={actionError} busy={acting} onRetry={pending.length ? retry : undefined} />}
          </Card>

          <div className="dd-histories">
            <OwnerHistory device={device} stints={history.stints} loading={history.loading} />
            <SpecHistory changes={history.changes} loading={history.loading} />
          </div>
          {history.error && <ErrorBanner message={history.error} onRetry={history.reload} />}

          <div className="dd-summary">
            <span className={`dd-risk rg-risk-${String(device.riskLevel).toLowerCase()}`}>
              {device.riskLevel ?? 'Unknown'}
              {typeof device.riskScore === 'number' && (
                <span className="dd-risk-score">{device.riskScore}</span>
              )}
            </span>
            <span className="dd-scanned">
              Scanned {formatMYT(device.scannedOn, 'datetime12')}
              <span className="dd-scanned-zone"> Malaysia time</span>
            </span>
            <label className="dd-toggle">
              <input
                type="checkbox"
                checked={showEmpty}
                onChange={(event) => setShowEmpty(event.target.checked)}
              />
              Show the fields the scan left blank
            </label>
            <label className="dd-toggle dd-toggle-tones">
              <input
                type="checkbox"
                checked={showTones}
                onChange={(event) => setShowTones(event.target.checked)}
              />
              Colour the risks red and the healthy values green
            </label>
          </div>

          <Card className="dd-fit">
            <h2 className="dd-group-title">
              Fit for the work
              <span className="dd-group-hint">{device.personaBlurb}</span>
            </h2>

            <div className="dd-fit-head">
              <span className={`dd-risk rg-risk-${FIT_TONE[device.fitStatus] ?? 'unknown'}`}>
                {device.fitStatus ?? 'Unknown'}
              </span>
              <span className="dd-fit-persona">{device.personaLabel}</span>
            </div>

            <ul className="dd-fit-reasons">
              {(device.fitReasons ?? []).map((reason) => (
                <li key={reason}>{reason}</li>
              ))}
            </ul>

            <dl className="dd-fit-facts">
              <div>
                <dt>Action</dt>
                <dd>{device.actionRequired ?? '—'}</dd>
              </div>
              <div>
                <dt>Suggested form factor</dt>
                <dd>
                  {device.suggestedFormFactor ?? '—'}
                  <span className="dd-fit-note">{device.formFactorNote}</span>
                </dd>
              </div>
              <div>
                <dt>Office licence</dt>
                <dd>
                  {device.licenseStatus ?? '—'}
                  <span className="dd-fit-note">{device.licenseNote}</span>
                </dd>
              </div>
              <div>
                <dt>Server link</dt>
                <dd>
                  {device.serverDependent ? device.networkRisk : 'Not server-bound'}
                  <span className="dd-fit-note">{device.networkNote}</span>
                </dd>
              </div>
            </dl>
          </Card>

          <div className="dd-groups">
            {groups.map((group) => (
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

          {device[RAW_REPORT_KEY] && (
            <Card className="dd-raw">
              <button
                type="button"
                className="dd-raw-toggle"
                onClick={() => setShowRaw((open) => !open)}
                aria-expanded={showRaw}
              >
                {showRaw ? 'Hide' : 'Show'} the scan report
                {device.sourceFileName && (
                  <span className="dd-raw-file">{formatScalar(device.sourceFileName)}</span>
                )}
              </button>
              {showRaw && <pre className="dd-raw-text">{device[RAW_REPORT_KEY]}</pre>}
            </Card>
          )}
        </>
      )}
    </AppShell>
  );
}
