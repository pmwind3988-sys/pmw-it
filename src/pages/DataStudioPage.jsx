import { useMemo, useRef, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { Inbox, AlertTriangle, RefreshCw } from '../components/ui/Icons';
import { DataStudioProvider } from '../features/datastudio/DataStudioContext';
import { useDataStudio } from '../features/datastudio/useDataStudio';
import CleanReview from '../features/datastudio/clean/CleanReview';

const ACCEPT = '.xlsx,.xlsm,.csv';

const TYPE_OPTIONS = ['numeric', 'categorical', 'boolean', 'date', 'datetime', 'text', 'identifier'];
const ROLE_OPTIONS = ['measure', 'dimension', 'temporal', 'ignored'];

function percent(ratio) {
  return `${Math.round((ratio ?? 0) * 100)}%`;
}

function DropStage() {
  const { importFile, error, fileName } = useDataStudio();
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  // Entering a child fires dragleave on the parent, so a plain boolean
  // would drop the highlight while the pointer is still inside. Same
  // enter/leave counting the device import uses.
  const depth = useRef(0);

  const stop = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  return (
    <>
      {error && <ErrorBanner message={error} />}
      <Card className="ds-drop-card">
        <div
          className={`ds-drop${dragging ? ' ds-drop-active' : ''}`}
          onDragEnter={(e) => { stop(e); depth.current += 1; setDragging(true); }}
          onDragOver={stop}
          onDragLeave={(e) => {
            stop(e);
            depth.current = Math.max(0, depth.current - 1);
            if (depth.current === 0) setDragging(false);
          }}
          onDrop={(e) => {
            stop(e);
            depth.current = 0;
            setDragging(false);
            const file = e.dataTransfer?.files?.[0];
            if (file) importFile(file);
          }}
        >
          <Inbox size={28} className="ds-drop-icon" />
          <p className="ds-drop-title">Drop a spreadsheet here</p>
          <p className="ds-drop-hint">
            An <code>.xlsx</code>, <code>.xlsm</code> or <code>.csv</code> file. It is read in
            your browser and never uploaded anywhere.
          </p>
          <Button variant="secondary" onClick={() => inputRef.current?.click()}>
            Choose a file
          </Button>
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPT}
            className="ds-drop-input"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) importFile(file);
              e.target.value = '';
            }}
          />
          {fileName && !error && <p className="ds-drop-hint">Last file: {fileName}</p>}
        </div>
      </Card>
    </>
  );
}

function ParsingStage() {
  const { progress, fileName } = useDataStudio();
  const pct = Math.max(0, Math.min(100, progress.pct ?? 0));
  return (
    <Card className="ds-progress-card">
      <p className="ds-progress-file">{fileName}</p>
      <div
        className="ds-progress-track"
        role="progressbar"
        aria-valuenow={pct}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-label={progress.stage || 'Importing'}
      >
        <div className="ds-progress-bar" style={{ width: `${pct}%` }} />
      </div>
      <p className="ds-progress-stage">{progress.stage || 'Working...'}</p>
    </Card>
  );
}

function ColumnRow({ column }) {
  const { overrideColumn, overrides } = useDataStudio();
  const override = overrides[column.name] ?? {};

  return (
    <tr>
      <td className="ds-col-name">
        {column.name}
        {column.overridden && <span className="ds-badge ds-badge-edit">edited</span>}
      </td>
      <td>
        <select
          className="ds-select"
          aria-label={`Type for ${column.name}`}
          value={column.type}
          onChange={(e) => overrideColumn(column.name, {
            // Changing the type re-derives the role unless the user has
            // separately pinned one, so picking "categorical" does not
            // leave a column still claiming to be a measure.
            type: e.target.value,
            role: override.role,
          })}
        >
          {TYPE_OPTIONS.map((t) => <option key={t} value={t}>{t}</option>)}
          {column.type === 'empty' && <option value="empty">empty</option>}
        </select>
      </td>
      <td>
        <select
          className="ds-select"
          aria-label={`Role for ${column.name}`}
          value={column.role}
          onChange={(e) => overrideColumn(column.name, {
            type: override.type ?? column.type,
            role: e.target.value,
          })}
        >
          {ROLE_OPTIONS.map((r) => <option key={r} value={r}>{r}</option>)}
        </select>
      </td>
      <td className="ds-num">{percent(column.nonNullRatio)}</td>
      <td className="ds-num">{column.distinctCount}</td>
      <td>
        {column.casualtyCount > 0 ? (
          <span className="ds-casualty">
            <AlertTriangle size={13} />
            {column.casualtyCount} value{column.casualtyCount === 1 ? '' : 's'} did not fit
            {column.casualties.length > 0 && (
              <span className="ds-casualty-eg">
                e.g. {column.casualties.slice(0, 3).map((c) => `"${c}"`).join(', ')}
              </span>
            )}
          </span>
        ) : (
          <span className="ds-ok">clean</span>
        )}
      </td>
    </tr>
  );
}

function ProfileStage() {
  const {
    profile, grid, sheets, activeSheet, selectSheet,
    headerIndex, headerCandidates, setHeaderIndex, reset, setStage,
  } = useDataStudio();

  const summary = useMemo(() => {
    const counts = { measure: 0, dimension: 0, temporal: 0, ignored: 0 };
    for (const c of profile?.columns ?? []) counts[c.role] = (counts[c.role] ?? 0) + 1;
    return counts;
  }, [profile]);

  if (!profile || !grid) return <EmptyState>Nothing imported yet.</EmptyState>;

  return (
    <>
      <div className="ds-toolbar">
        {sheets.length > 1 && (
          <label className="ds-field">
            <span>Sheet</span>
            <select
              className="ds-select"
              value={activeSheet}
              onChange={(e) => selectSheet(e.target.value)}
            >
              {sheets.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
        )}

        <label className="ds-field">
          <span>Header row</span>
          <select
            className="ds-select"
            value={headerIndex}
            onChange={(e) => setHeaderIndex(Number(e.target.value))}
          >
            {headerCandidates.map((row, i) => (
              <option key={i} value={i}>
                {`Row ${i + 1} - ${row.filter(Boolean).slice(0, 4).join(', ') || '(blank)'}`}
              </option>
            ))}
          </select>
        </label>

        <span className="ds-toolbar-spacer" />
        <span className="ds-summary">
          {`${profile.rowCount.toLocaleString()} rows · ${profile.columns.length} columns · `}
          {`${summary.measure} measures · ${summary.dimension} dimensions · `}
          {`${summary.temporal} temporal`}
        </span>
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reset}>
          Start over
        </Button>
        <Button size="sm" onClick={() => setStage('cleaning')}>
          Continue
        </Button>
      </div>

      <Card className="ds-table-card">
        <div className="ds-table-scroll">
          <table className="ds-table">
            <thead>
              <tr>
                <th>Column</th>
                <th>Type</th>
                <th>Role</th>
                <th className="ds-num">Filled</th>
                <th className="ds-num">Distinct</th>
                <th>Fit</th>
              </tr>
            </thead>
            <tbody>
              {profile.columns.map((column) => (
                <ColumnRow key={column.name} column={column} />
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </>
  );
}

function DataStudioBody() {
  const { stage } = useDataStudio();
  if (stage === 'parsing') return <ParsingStage />;
  if (stage === 'idle') return <DropStage />;
  if (stage === 'cleaning') return <CleanReview />;
  return <ProfileStage />;
}

export default function DataStudioPage() {
  return (
    <DataStudioProvider>
      <AppShell
        title="Data Studio"
        subtitle="Import a spreadsheet and chart it, without leaving the browser."
      >
        <DataStudioBody />
      </AppShell>
    </DataStudioProvider>
  );
}
