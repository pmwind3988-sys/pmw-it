import { useEffect, useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';

/**
 * The verdict shares the risk palette rather than a second one: red is "go and
 * look", amber "put it on the list", green "leave it alone".
 */
const FIT_TONE = {
  Critical: 'critical',
  'Needs Attention': 'watch',
  Optimal: 'ok',
};
import Button from '../components/ui/Button';
import { ArrowLeft, RefreshCw } from '../components/ui/Icons';
import { useDevices } from '../features/devices/useDevices';
import { groupsFor, RAW_REPORT_KEY } from '../features/devices/fieldGroups';
import { formatScalar } from '../features/devices/formatValue';
import ValueCell from '../features/devices/ui/ValueCell';
import { toneForField, toneForEntry, hasEntryTones } from '../features/devices/fieldTone';
import { formatMYT } from '../utils/malaysiaTime';

/** Remembered per browser: somebody who turns the colouring off is not asked
 *  to turn it off again on the next machine they open. */
const TONE_KEY = 'deviceValueTones';

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
