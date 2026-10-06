import { useCallback, useMemo, useRef, useState } from 'react';
import { useSearchParams, useNavigate } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { gradeCssVars } from '../features/devices/standards/gradeColors';
import StatCard from '../components/ui/StatCard';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import {
  Laptop, AlertTriangle, ShieldCheck, MemoryStick, Clock, RefreshCw, Tag, WifiOff, Archive,
} from '../components/ui/Icons';
import { useSharePointToken } from '../hooks/useRequests';
import DropZone from '../features/devices/ui/DropZone';
import ReviewGrid from '../features/devices/ui/ReviewGrid';
import SaveProgress from '../features/devices/ui/SaveProgress';
import DeviceTable from '../features/devices/ui/DeviceTable';
import DeviceCharts from '../features/devices/ui/DeviceCharts';
import DepartmentHeatmap from '../features/devices/ui/DepartmentHeatmap';
import Leaderboards from '../features/devices/ui/Leaderboards';
import DeviceMap from '../features/devices/ui/DeviceMap';
import { importFiles, mergeImports } from '../features/devices/importFiles';
import { issuesFor, sortForReview } from '../features/devices/reviewIssues';
import { useDevices } from '../features/devices/useDevices';
import { fleetSummary, complianceSummary } from '../features/devices/stats/deviceStats';
import { labelOf, IN_FLEET } from '../features/devices/deviceFilters';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';
import { updateDevice, deleteDevice, deleteDevices } from '../features/devices/sharepoint/updateDevice';
import { provisionLists } from '../features/devices/sharepoint/provisionLists';
import { matchIncoming, noticeFor } from '../features/devices/lifecycle/matchIncoming';
import { replacementsFor, unanswered } from '../features/devices/lifecycle/replacements';
import { locationsIn, cleanLocation } from '../features/devices/map/locations';
import { fileNameLocations } from '../features/devices/map/fileNameLocations';
import { inFleet, statusOf, STATUSES, RETIRED, SPARE } from '../features/devices/lifecycle/status';
import { performLifecycle } from '../features/devices/sharepoint/writeLifecycle';
import { ACTIONS } from '../features/devices/lifecycle/planLifecycle';
import { mapHref } from '../features/devices/map/mapLinks';
import { PLACES } from '../features/devices/map/zones';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const IDLE_SAVE = {
  phase: 'starting', done: 0, total: 0, results: null, error: null,
  changeCount: 0, unchanged: 0, stintFailures: 0, skipped: [],
};

/** Everything in the query string that is a filter rather than a view switch. */
const FILTER_KEYS = [
  'risk', 'attention', 'type', 'department', 'os', 'av',
  'storage', 'ram', 'cpu', 'windows', 'stale', 'q',
  'part', 'critical', 'persona', 'license', 'server', 'formfit', 'status', 'location',
];

export default function DevicesPage() {
  const getToken = useSharePointToken();
  const [params, setParams] = useSearchParams();
  const { devices: saved, loading, error, reload, standards } = useDevices();

  const view = params.get('view') ?? 'map';

  const [stage, setStage] = useState('drop');
  const [parsed, setParsed] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  // Edits are held apart from the parsed records so that a re-parse or a
  // "start over" discards them cleanly, and the raw record still matches the
  // file it came from.
  const [edits, setEdits] = useState({});
  const [excluded, setExcluded] = useState(new Set());
  const [answers, setAnswers] = useState({});
  const [save, setSave] = useState(IDLE_SAVE);
  const [rowBusy, setRowBusy] = useState(false);
  const [rowError, setRowError] = useState('');
  const columnsChecked = useRef(false);

  const merged = useMemo(
    () => parsed.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) })),
    [parsed, edits],
  );

  const matches = useMemo(() => matchIncoming(merged, saved), [merged, saved]);
  const prompts = useMemo(() => replacementsFor(matches, saved), [matches, saved]);
  const notices = useMemo(() => new Map(
    matches.map((match) => [match.device.sourceFileName, noticeFor(match)]).filter(([, text]) => text),
  ), [matches]);
  const waiting = unanswered(prompts, answers, excluded);

  const filters = useMemo(
    () => Object.fromEntries(FILTER_KEYS.map((key) => [key, params.get(key) ?? ''])),
    [params],
  );

  const navigate = useNavigate();
  // Figures count the machines people are working on. A retired laptop
  // reported as a critical risk would be a figure lying about the fleet.
  const fleet = useMemo(() => saved.filter(inFleet), [saved]);
  const spareCount = useMemo(() => saved.filter((d) => statusOf(d) === SPARE).length, [saved]);
  // The register hides retired machines unless a status filter asks for them.
  const registerRows = useMemo(
    () => (filters.status ? saved : saved.filter((d) => statusOf(d) !== RETIRED)),
    [saved, filters.status],
  );
  const locationOptions = useMemo(() => locationsIn(saved), [saved]);

  // The dashboard reads one department at a time when asked to. It shares the
  // register's `department` key, so a scope chosen here survives the jump into
  // the rows behind any card.
  const department = params.get('department') ?? '';

  const departments = useMemo(() => {
    const names = new Set(fleet.map((device) => labelOf(device.department)));
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [fleet]);

  const scoped = useMemo(
    () => (department
      ? fleet.filter((device) => labelOf(device.department) === department)
      : fleet),
    [fleet, department],
  );

  const summary = useMemo(() => fleetSummary(scoped), [scoped]);
  const compliance = useMemo(() => complianceSummary(scoped), [scoped]);

  const flagged = merged.filter((device) => issuesFor(device).length > 0).length;
  const included = merged.filter((device) => !excluded.has(device.sourceFileName)).length;

  const setParam = useCallback((key, value) => {
    setParams((current) => {
      const next = new URLSearchParams(current);
      if (value) next.set(key, value);
      else next.delete(key);
      return next;
    });
  }, [setParams]);

  /** Open the register, optionally filtered. `key` null means "no filter". */
  const openRegister = useCallback((key, value) => {
    setParams((current) => {
      const next = new URLSearchParams(current);
      next.set('view', 'register');
      // Always default to In use + In repair, but let an explicit status key override it
      if (key !== 'status') next.set('status', IN_FLEET);
      if (!key || key === 'view') return next;
      if (value) {
        next.set(key, value);
      } else {
        next.delete(key);
      }
      return next;
    });
  }, [setParams]);

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const incoming = await importFiles(files, { knownLocations: fileNameLocations(saved) });
      // A second drop adds to the review rather than starting it over, so a
      // batch that arrives in several goes still ends up as one save. Edits
      // already made are keyed by file name and survive untouched.
      const result = parsed.length
        ? mergeImports({ devices: parsed, rejected: [] }, incoming)
        : incoming;
      // Sorted once per drop: the grid must not reorder while somebody edits it.
      const asking = new Set(replacementsFor(matchIncoming(result.devices, saved), saved)
        .map((prompt) => prompt.sourceFileName));
      setParsed(sortForReview(result.devices, asking));
      setRejected(result.rejected);
      if (result.devices.length) setStage('review');
    } finally {
      setBusy(false);
    }
  }, [parsed, saved]);

  const handleChange = (id, key, value) =>
    setEdits((current) => ({ ...current, [id]: { ...(current[id] ?? {}), [key]: value } }));

  const handleToggleRow = (id) =>
    setExcluded((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const resetImport = () => {
    setParsed([]);
    setRejected([]);
    setEdits({});
    setExcluded(new Set());
    setAnswers({});
    setSave(IDLE_SAVE);
    setStage('drop');
  };

  const handleSave = async (onlyNames) => {
    const toSave = merged.filter(
      (device) =>
        !excluded.has(device.sourceFileName)
        && (!onlyNames || onlyNames.includes(device.computerName)),
    );

    setStage('save');
    setSave({ ...IDLE_SAVE, total: toSave.length });

    try {
      const tokenRes = await getToken();
      const outcome = await syncDevices({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        devices: toSave,
        answers,
        changedBy: tokenRes.account?.username ?? '',
        onProgress: ({ phase, done, total }) =>
          setSave((current) => ({ ...current, phase, done, total })),
      });
      setSave((current) => ({
        ...current,
        results: outcome.results,
        changeCount: outcome.changeCount,
        unchanged: outcome.unchanged,
        stintFailures: outcome.stintFailures,
        skipped: outcome.skipped,
      }));
      reload();
    } catch (failure) {
      setSave((current) => ({ ...current, error: failure.message }));
    }
  };

  /** One row edited or removed in the register, rather than a whole import. */
  const runRowAction = async (action) => {
    setRowBusy(true);
    setRowError('');
    try {
      const tokenRes = await getToken();

      // Once per visit, and cheap when there is nothing to do: an edit writes
      // columns a save would have created, so the register must not depend on
      // somebody having run an import first.
      if (!columnsChecked.current) {
        await provisionLists(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        columnsChecked.current = true;
      }

      try {
        await action(tokenRes);
      } finally {
        // The register re-reads whatever the action managed to do, failure or
        // not: removing twelve machines and getting eleven has still changed
        // SharePoint, and a table left showing all twelve would be lying.
        reload();
      }
    } catch (failure) {
      setRowError(failure.message);
    } finally {
      setRowBusy(false);
    }
  };

  const OWNERSHIP = ['owner', 'location', 'department'];
  const trimmed = (value) => String(value ?? '').trim();

  /**
   * An owner, location or department typed into the register is a change of
   * hands, so it goes through the same write as the machine page's Change
   * owner -- the register and the owner history cannot disagree. Clearing the
   * owner still goes the old way: a blank hands the field back to the scan.
   */
  const handleRowSave = (device, edits) => runRowAction(async (tokenRes) => {
    const moved = OWNERSHIP.some((key) => {
      if (!(key in edits)) return false;
      if (key === 'location') return cleanLocation(edits[key]) !== cleanLocation(device[key]);
      // A capitalisation-only edit is not a change of hands; the lifecycle
      // write refuses it, so it goes through the plain edit below.
      return trimmed(edits[key]).toUpperCase() !== trimmed(device[key]).toUpperCase();
    });
    const owner = trimmed('owner' in edits ? edits.owner : device.owner);
    let existing = device;
    let rest = edits;
    // logFailed is deliberately not an error here: the move happened.

    if (moved && owner) {
      // A machine in the stash or the graveyard being given an owner is an
      // assignment: the plain edit would leave a Spare machine with an owner.
      const { plan, pendingStints } = await performLifecycle({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        deviceId: device.id,
        action: inFleet(device) ? ACTIONS.CHANGE_OWNER : ACTIONS.ASSIGN,
        input: {
          owner,
          location: 'location' in edits ? edits.location : device.location,
          department: 'department' in edits ? edits.department : device.department,
        },
        expectedStatus: statusOf(device),
        recordedBy: tokenRes.account?.username ?? '',
      });
      // The remaining edit must start from the manual list the change just
      // wrote, or saving the device type would drop owner from it again.
      existing = { ...device, ...plan.fields };
      rest = Object.fromEntries(Object.entries(edits).filter(([key]) => !OWNERSHIP.includes(key)));
      if (pendingStints.length) {
        throw new Error('The owner changed, but its history could not be written. Open the machine to retry.');
      }
    }

    if (Object.keys(rest).length) {
      await updateDevice({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        existing,
        edits: rest,
        changedBy: tokenRes.account?.username ?? '',
      });
    }
  });

  const handleRowDelete = (device) => runRowAction((tokenRes) => deleteDevice({
    siteUrl: SHAREPOINT_SITE_URL,
    token: tokenRes.accessToken,
    device,
    changedBy: tokenRes.account?.username ?? '',
  }));

  /**
   * A ticked selection removed in one go. `deleteDevices` reports rather than
   * throws, so a machine that would not go has to be named here — the rest of
   * the selection has already gone, and a silent partial removal would leave
   * somebody believing the register is emptier than it is.
   */
  const handleRowDeleteMany = (chosen, onProgress) =>
    runRowAction(async (tokenRes) => {
      const outcome = await deleteDevices({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        devices: chosen,
        changedBy: tokenRes.account?.username ?? '',
        onProgress,
      });

      if (outcome.failures.length) {
        const names = outcome.failures.map((failure) => failure.computerName).join(', ');
        throw new Error(
          `Removed ${outcome.removed.length} of ${chosen.length}. `
          + `Could not remove ${names}: ${outcome.failures[0].error}`,
        );
      }
    });

  const rejectedList = rejected.length > 0 && (
    <ul className="dz-rejected">
      {rejected.map((item) => (
        <li key={item.fileName}>
          <strong>{item.fileName}</strong> — {item.reason}
        </li>
      ))}
    </ul>
  );

  const scopePicker = view === 'dashboard' && departments.length > 0 && (
    <label className="dv-scope">
      <span>Department</span>
      <select
        value={department}
        onChange={(event) => setParam('department', event.target.value)}
      >
        <option value="">All departments</option>
        {departments.map((name) => (
          <option key={name} value={name}>{name}</option>
        ))}
      </select>
    </label>
  );

  const tabs = (
    <div className="dv-tabs" role="tablist">
      {[
        ['map', 'Map'],
        ['dashboard', 'Dashboard'],
        ['register', 'Register'],
        ['import', 'Import'],
      ].map(([key, label]) => (
        <button
          type="button"
          role="tab"
          key={key}
          aria-selected={view === key}
          className={`dv-tab${view === key ? ' dv-tab-active' : ''}`}
          onClick={() => setParams((current) => {
            const next = new URLSearchParams(current);
            next.set('view', key);
            ['location', 'place'].forEach((k) => next.delete(k));
            if (view === 'map') next.delete('department');
            return next;
          })}
        >
          {label}
        </button>
      ))}
      {scopePicker}
    </div>
  );

  return (
    <AppShell
      title="Device list"
      subtitle={department
        ? `${department} — what every machine has, what needs attention, and what is getting old`
        : 'What every machine has, what needs attention, and what is getting old'}
      actions={(
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reload} loading={loading}>
          Refresh
        </Button>
      )}
    >
      <div className="dv-graded" style={gradeCssVars(standards.standard.colors)}>
        {(standards.note || standards.error) && (
          <p className="dv-standard-note" role="status">{standards.note || standards.error}</p>
        )}

        {tabs}

      {error && <ErrorBanner message={error} onRetry={reload} />}

      {view === 'map' && <DeviceMap devices={saved} loading={loading} params={params} standard={standards.standard} />}

      {view === 'dashboard' && (
        <>
          <div className="stat-grid">
            <StatCard
              icon={Laptop}
              label="Devices"
              value={summary.total}
              loading={loading}
              onClick={() => openRegister(null, null)}
            />
            <StatCard
              icon={AlertTriangle}
              label="Need attention"
              value={summary.needsAttention}
              color="var(--it-danger)"
              loading={loading}
              onClick={() => openRegister('attention', '1')}
            />
            <StatCard
              icon={ShieldCheck}
              label="Unsupported OS"
              value={summary.unsupportedOs}
              color="var(--it-danger)"
              loading={loading}
              onClick={() => openRegister('os', 'Unsupported')}
            />
            <StatCard
              icon={ShieldCheck}
              label="Unprotected"
              value={summary.unprotected}
              color="var(--it-accent)"
              loading={loading}
              onClick={() => openRegister('av', 'Unprotected')}
            />
            <StatCard
              icon={MemoryStick}
              label="Average RAM"
              value={summary.avgRamGB ?? '—'}
              unit="GB"
              loading={loading}
            />
            <StatCard
              icon={Clock}
              label="Stale scans"
              value={summary.staleScans}
              loading={loading}
              onClick={() => openRegister('stale', '1')}
            />
            <StatCard
              icon={AlertTriangle}
              label="Machines with a critical part"
              value={compliance.criticalPct ?? '—'}
              unit={`% · ${compliance.critical} machines`}
              color="var(--it-danger)"
              loading={loading}
              onClick={() => openRegister('critical', '1')}
            />
            <StatCard
              icon={ShieldCheck}
              label="Office compliance"
              value={compliance.complianceRate ?? '—'}
              unit="%"
              color={compliance.unlicensed ? 'var(--it-danger)' : 'var(--it-good)'}
              loading={loading}
              onClick={() => openRegister('license', 'Unlicensed')}
            />
            <StatCard
              icon={Archive}
              label="In IT Stash"
              value={spareCount}
              loading={loading}
              onClick={() => navigate(mapHref({ place: PLACES.STASH }))}
            />
            <StatCard
              icon={WifiOff}
              label="Server over Wi-Fi"
              value={compliance.networkBottlenecks}
              color="var(--it-danger)"
              loading={loading}
              onClick={() => openRegister('server', 'Bottleneck')}
            />
            <StatCard
              icon={Tag}
              label="Form factor mismatch"
              value={compliance.mismatchedFormFactor}
              loading={loading}
              onClick={() => openRegister('formfit', '1')}
            />
          </div>

          {!loading && scoped.length === 0 ? (
            <Card>
              <EmptyState>
                {saved.length === 0
                  ? 'Nothing in the register yet. Open the Import tab and drop your scan reports.'
                  : `No devices are recorded against ${department}.`}
              </EmptyState>
            </Card>
          ) : (
            <>
              <DepartmentHeatmap
                devices={scoped}
                onSelect={(name, part) => {
                  setParam('department', name);
                  openRegister('part', `${part}:Critical`);
                }}
              />
              <DeviceCharts devices={scoped} onFilter={openRegister} />
              <Leaderboards devices={scoped} />
            </>
          )}
        </>
      )}

      {view === 'register' && (
        <>
          {rowError && <ErrorBanner message={rowError} onRetry={() => setRowError('')} />}
          <div className="dv-register-scope">
            <label className="dv-scope">
              <span>Location</span>
              <select value={filters.location} onChange={(event) => setParam('location', event.target.value)}>
                <option value="">All locations</option>
                {locationOptions.map((code) => <option key={code} value={code}>{code}</option>)}
                <option value="Unassigned">No location yet</option>
              </select>
            </label>
            <label className="dv-scope">
              <span>Status</span>
              <select value={filters.status} onChange={(event) => setParam('status', event.target.value)}>
                <option value="">All but retired</option>
                <option value={IN_FLEET}>In use or in repair</option>
                {STATUSES.map((status) => <option key={status} value={status}>{status}</option>)}
              </select>
            </label>
            {!filters.status && (
              <button type="button" className="dv-linkish" onClick={() => setParam('status', RETIRED)}>
                Retired machines are hidden. Show them
              </button>
            )}
          </div>
          <DeviceTable
            devices={registerRows}
            filters={filters}
            onFilterChange={setParam}
            onSave={handleRowSave}
            onDelete={handleRowDelete}
            onDeleteMany={handleRowDeleteMany}
            busy={rowBusy}
          />
        </>
      )}

      {view === 'import' && stage === 'drop' && (
        <Card className="dz-card">
          <DropZone onFiles={handleFiles} busy={busy} />
          {rejectedList}
        </Card>
      )}

      {view === 'import' && stage === 'review' && (
        <Card className="rg-card">
          <div className="review-head">
            <p className="review-summary">
              {included} of {merged.length} selected
              {flagged > 0 && <span className="review-flagged"> · {flagged} need attention</span>}
              {waiting > 0 && <span className="review-flagged"> · {waiting} replacement{waiting === 1 ? '' : 's'} to answer</span>}
            </p>
            <div className="review-actions">
              <Button variant="secondary" size="sm" onClick={resetImport}>Start over</Button>
              <Button size="sm" disabled={included === 0 || waiting > 0} onClick={() => handleSave(null)}>
                {waiting > 0
                  ? `Save — ${waiting} replacement${waiting === 1 ? '' : 's'} need${waiting === 1 ? 's' : ''} an answer`
                  : `Save ${included} to SharePoint`}
              </Button>
            </div>
          </div>

          {merged.length === 0 ? (
            <EmptyState>Nothing to review.</EmptyState>
          ) : (
            <ReviewGrid
              devices={merged}
              excluded={excluded}
              onChange={handleChange}
              onToggleRow={handleToggleRow}
              prompts={prompts}
              answers={answers}
              notices={notices}
              onAnswer={(key, value) => setAnswers((current) => ({ ...current, [key]: value }))}
            />
          )}

          <DropZone onFiles={handleFiles} busy={busy} compact />

          {rejectedList}
        </Card>
      )}

      {view === 'import' && stage === 'save' && (
        <Card>
          <SaveProgress state={save} onRetry={handleSave} onDone={resetImport} />
        </Card>
      )}
      </div>
    </AppShell>
  );
}
