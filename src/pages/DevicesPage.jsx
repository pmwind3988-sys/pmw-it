import { useCallback, useMemo, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { useSharePointToken } from '../hooks/useRequests';
import DropZone from '../features/devices/ui/DropZone';
import ReviewGrid from '../features/devices/ui/ReviewGrid';
import SaveProgress from '../features/devices/ui/SaveProgress';
import { importFiles } from '../features/devices/importFiles';
import { issuesFor } from '../features/devices/reviewIssues';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const IDLE_SAVE = {
  done: 0, total: 0, results: null, error: null, changeCount: 0, unchanged: 0,
};

export default function DevicesPage() {
  const getToken = useSharePointToken();

  const [stage, setStage] = useState('drop');
  const [devices, setDevices] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  // Edits are held apart from the parsed records so that a re-parse or a
  // "start over" discards them cleanly, and the raw record still matches the
  // file it came from.
  const [edits, setEdits] = useState({});
  const [excluded, setExcluded] = useState(new Set());
  const [save, setSave] = useState(IDLE_SAVE);

  const merged = useMemo(
    () => devices.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) })),
    [devices, edits],
  );

  const flagged = merged.filter((device) => issuesFor(device).length > 0).length;
  const included = merged.filter((device) => !excluded.has(device.sourceFileName)).length;

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const result = await importFiles(files);
      setDevices(result.devices);
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

  const reset = () => {
    setDevices([]);
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
    } catch (error) {
      setSave((current) => ({ ...current, error: error.message }));
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

  return (
    <AppShell
      title="Device list"
      subtitle="Import machine scan reports and keep the fleet register current"
    >
      {stage === 'drop' && (
        <Card className="dz-card">
          <DropZone onFiles={handleFiles} busy={busy} />
          {rejectedList}
        </Card>
      )}

      {stage === 'review' && (
        <Card>
          <div className="review-head">
            <p className="review-summary">
              {included} of {merged.length} selected
              {flagged > 0 && <span className="review-flagged"> · {flagged} need attention</span>}
            </p>
            <div className="review-actions">
              <Button variant="secondary" size="sm" onClick={reset}>Start over</Button>
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

      {stage === 'save' && (
        <Card>
          <SaveProgress
            state={save}
            onRetry={(names) => handleSave(names)}
            onDone={reset}
          />
        </Card>
      )}
    </AppShell>
  );
}
