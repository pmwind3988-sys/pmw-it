import { useCallback, useMemo, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import DropZone from '../features/devices/ui/DropZone';
import ReviewGrid from '../features/devices/ui/ReviewGrid';
import { importFiles } from '../features/devices/importFiles';
import { issuesFor } from '../features/devices/reviewIssues';

export default function DevicesPage() {
  const [stage, setStage] = useState('drop');
  const [devices, setDevices] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  // Edits are held apart from the parsed records so that a re-parse or a
  // "start over" discards them cleanly, and the raw record still matches the
  // file it came from.
  const [edits, setEdits] = useState({});
  const [excluded, setExcluded] = useState(new Set());

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
    setStage('drop');
  };

  return (
    <AppShell
      title="Device list"
      subtitle="Import machine scan reports and keep the fleet register current"
    >
      {stage === 'drop' && (
        <Card className="dz-card">
          <DropZone onFiles={handleFiles} busy={busy} />
          {rejected.length > 0 && (
            <ul className="dz-rejected">
              {rejected.map((item) => (
                <li key={item.fileName}>
                  <strong>{item.fileName}</strong> — {item.reason}
                </li>
              ))}
            </ul>
          )}
        </Card>
      )}

      {stage === 'review' && (
        <Card>
          <div className="review-head">
            <p className="review-summary">
              {included} of {merged.length} selected
              {flagged > 0 && <span className="review-flagged"> · {flagged} need attention</span>}
            </p>
            <Button variant="secondary" size="sm" onClick={reset}>Start over</Button>
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

          {rejected.length > 0 && (
            <ul className="dz-rejected">
              {rejected.map((item) => (
                <li key={item.fileName}>
                  <strong>{item.fileName}</strong> — {item.reason}
                </li>
              ))}
            </ul>
          )}
        </Card>
      )}
    </AppShell>
  );
}
