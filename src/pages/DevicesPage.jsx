import { useCallback, useMemo, useRef, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import StatCard from '../components/ui/StatCard';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import {
  Laptop, AlertTriangle, ShieldCheck, MemoryStick, Clock, RefreshCw,
} from '../components/ui/Icons';
import { useSharePointToken } from '../hooks/useRequests';
import DropZone from '../features/devices/ui/DropZone';
import ReviewGrid from '../features/devices/ui/ReviewGrid';
import SaveProgress from '../features/devices/ui/SaveProgress';
import DeviceTable from '../features/devices/ui/DeviceTable';
import DeviceCharts from '../features/devices/ui/DeviceCharts';
import Leaderboards from '../features/devices/ui/Leaderboards';
import { importFiles, mergeImports } from '../features/devices/importFiles';
import { issuesFor, sortForReview } from '../features/devices/reviewIssues';
import { useDevices } from '../features/devices/useDevices';
import { fleetSummary } from '../features/devices/stats/deviceStats';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';
import { updateDevice, deleteDevice } from '../features/devices/sharepoint/updateDevice';
import { provisionLists } from '../features/devices/sharepoint/provisionLists';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const IDLE_SAVE = {
  phase: 'starting', done: 0, total: 0, results: null, error: null,
  changeCount: 0, unchanged: 0,
};

/** Everything in the query string that is a filter rather than a view switch. */
const FILTER_KEYS = [
  'risk', 'type', 'department', 'os', 'av', 'storage', 'ram', 'cpu', 'windows', 'stale', 'q',
];

export default function DevicesPage() {
  const getToken = useSharePointToken();
  const [params, setParams] = useSearchParams();
  const { devices: saved, loading, error, reload } = useDevices();

  const view = params.get('view') ?? 'dashboard';

  const [stage, setStage] = useState('drop');
  const [parsed, setParsed] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  // Edits are held apart from the parsed records so that a re-parse or a
  // "start over" discards them cleanly, and the raw record still matches the
  // file it came from.
  const [edits, setEdits] = useState({});
  const [excluded, setExcluded] = useState(new Set());
  const [save, setSave] = useState(IDLE_SAVE);
  const [rowBusy, setRowBusy] = useState(false);
  const [rowError, setRowError] = useState('');
  const columnsChecked = useRef(false);

  const merged = useMemo(
    () => parsed.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) })),
    [parsed, edits],
  );

  const filters = useMemo(
    () => Object.fromEntries(FILTER_KEYS.map((key) => [key, params.get(key) ?? ''])),
    [params],
  );

  // The dashboard reads one department at a time when asked to. It shares the
  // register's `department` key, so a scope chosen here survives the jump into
  // the rows behind any card.
  const department = params.get('department') ?? '';

  const departments = useMemo(() => {
    const names = new Set(saved.map((device) => device.department || 'Unassigned'));
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [saved]);

  const scoped = useMemo(
    () => (department
      ? saved.filter((device) => (device.department || 'Unassigned') === department)
      : saved),
    [saved, department],
  );

  const summary = useMemo(() => fleetSummary(scoped), [scoped]);

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
      if (!key || key === 'view') return next;
      if (value) next.set(key, value);
      else next.delete(key);
      return next;
    });
  }, [setParams]);

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const incoming = await importFiles(files);
      // A second drop adds to the review rather than starting it over, so a
      // batch that arrives in several goes still ends up as one save. Edits
      // already made are keyed by file name and survive untouched.
      const result = parsed.length
        ? mergeImports({ devices: parsed, rejected: [] }, incoming)
        : incoming;
      // Sorted once per drop: the grid must not reorder while somebody edits it.
      setParsed(sortForReview(result.devices));
      setRejected(result.rejected);
      if (result.devices.length) setStage('review');
    } finally {
      setBusy(false);
    }
  }, [parsed]);

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
        changedBy: tokenRes.account?.username ?? '',
        onProgress: ({ phase, done, total }) =>
          setSave((current) => ({ ...current, phase, done, total })),
      });
      setSave((current) => ({
        ...current,
        results: outcome.results,
        changeCount: outcome.changeCount,
        unchanged: outcome.unchanged,
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

      await action(tokenRes);
      reload();
    } catch (failure) {
      setRowError(failure.message);
    } finally {
      setRowBusy(false);
    }
  };

  const handleRowSave = (device, edits) => runRowAction((tokenRes) => updateDevice({
    siteUrl: SHAREPOINT_SITE_URL,
    token: tokenRes.accessToken,
    existing: device,
    edits,
    changedBy: tokenRes.account?.username ?? '',
  }));

  const handleRowDelete = (device) => runRowAction((tokenRes) => deleteDevice({
    siteUrl: SHAREPOINT_SITE_URL,
    token: tokenRes.accessToken,
    device,
    changedBy: tokenRes.account?.username ?? '',
  }));

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
          onClick={() => setParam('view', key)}
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
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reload} disabled={loading}>
          Refresh
        </Button>
      )}
    >
      {tabs}

      {error && <ErrorBanner message={error} onRetry={reload} />}

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
              onClick={() => openRegister('risk', 'Critical')}
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
              <DeviceCharts devices={scoped} onFilter={openRegister} />
              <Leaderboards devices={scoped} />
            </>
          )}
        </>
      )}

      {view === 'register' && (
        <>
          {rowError && <ErrorBanner message={rowError} onRetry={() => setRowError('')} />}
          <DeviceTable
            devices={saved}
            filters={filters}
            onFilterChange={setParam}
            onSave={handleRowSave}
            onDelete={handleRowDelete}
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
        <Card>
          <div className="review-head">
            <p className="review-summary">
              {included} of {merged.length} selected
              {flagged > 0 && <span className="review-flagged"> · {flagged} need attention</span>}
            </p>
            <div className="review-actions">
              <Button variant="secondary" size="sm" onClick={resetImport}>Start over</Button>
              <Button size="sm" disabled={included === 0} onClick={() => handleSave(null)}>
                Save {included} to SharePoint
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
    </AppShell>
  );
}
