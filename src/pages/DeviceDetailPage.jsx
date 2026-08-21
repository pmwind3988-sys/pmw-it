import { useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { ArrowLeft, RefreshCw } from '../components/ui/Icons';
import { useDevices } from '../features/devices/useDevices';
import { groupsFor, RAW_REPORT_KEY } from '../features/devices/fieldGroups';
import { formatScalar } from '../features/devices/formatValue';
import ValueCell from '../features/devices/ui/ValueCell';
import { formatMYT } from '../features/datastudio/time/malaysiaTime';

/** One machine, everything the scan read out of it, grouped by what it is. */
export default function DeviceDetailPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { devices, loading, error, reload } = useDevices();
  const [showEmpty, setShowEmpty] = useState(false);
  const [showRaw, setShowRaw] = useState(false);

  const device = useMemo(
    () => devices.find((row) => String(row.id) === String(id)),
    [devices, id],
  );

  const groups = useMemo(
    () => (device ? groupsFor(device, { includeEmpty: showEmpty }) : []),
    [device, showEmpty],
  );

  const manual = new Set(device?.manualFields ?? []);

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
          <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reload} disabled={loading}>
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
                <Link to="/devices?view=register">Back to the register</Link>
              </>
            )}
          </EmptyState>
        </Card>
      ) : (
        <>
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
          </div>

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
