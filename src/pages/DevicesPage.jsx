import { useCallback, useMemo, useState } from 'react';
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
import { importFiles } from '../features/devices/importFiles';
import { issuesFor } from '../features/devices/reviewIssues';
import { useDevices } from '../features/devices/useDevices';
import { fleetSummary } from '../features/devices/stats/deviceStats';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const IDLE_SAVE = {
  done: 0, total: 0, results: null, error: null, changeCount: 0, unchanged: 0,
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

  const merged = useMemo(
    () => parsed.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) })),
    [parsed, edits],
  );

  const filters = useMemo(
    () => Object.fromEntries(FILTER_KEYS.map((key) => [key, params.get(key) ?? ''])),
    [params],
  );

  const summary = useMemo(() => fleetSummary(saved), [saved]);

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

  const openRegister = useCallback((key, value) => {
    setParams((current) => {
      const next = new URLSearchParams(current);
      next.set('view', 'register');
      if (value) next.set(key, value);
      else next.delete(key);
      return next;
    });
  }, [setParams]);

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const result = await importFiles(files);
      setParsed(result.devices);
      setRejected(result.rejected);
      setEdits({});
      setExcluded(new Set());
      setStage(result.devices.length ? 'review' : 'drop');
    } finally {
      setBusy(false);
    }
  }, []);

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
        onProgress: (done, total) => setSave((current) => ({ ...current, done, total })),
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

  const rejectedList = rejected.length > 0 && (
    <ul className="dz-rejected">
      {rejected.map((item) => (
        <li key={item.fileName}>
          <strong>{item.fileName}</strong> — {item.reason}
        </li>
      ))}
    </ul>
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
    </div>
  );

  return (
    <AppShell
      title="Device list"
      subtitle="What every machine has, what needs attention, and what is getting old"
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
              onClick={() => openRegister('view', null)}
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

          {!loading && saved.length === 0 ? (
            <Card>
              <EmptyState>
                Nothing in the register yet. Open the Import tab and drop your scan reports.
              </EmptyState>
            </Card>
          ) : (
            <>
              <DeviceCharts devices={saved} onFilter={openRegister} />
              <Leaderboards devices={saved} />
            </>
          )}
        </>
      )}

      {view === 'register' && (
        <DeviceTable devices={saved} filters={filters} onFilterChange={setParam} />
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
