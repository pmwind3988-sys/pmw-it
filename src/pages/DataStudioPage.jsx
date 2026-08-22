import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { Inbox, AlertTriangle, RefreshCw, Plus, Trash2 } from '../components/ui/Icons';
import { DataStudioProvider } from '../features/datastudio/DataStudioContext';
import { useDataStudio } from '../features/datastudio/useDataStudio';
import CleanReview from '../features/datastudio/clean/CleanReview';
import CanvasGrid from '../features/datastudio/canvas/CanvasGrid';
import TileEditor from '../features/datastudio/canvas/TileEditor';
import FilterBar from '../features/datastudio/canvas/FilterBar';
import DatasetLibrary from '../features/datastudio/store/DatasetLibrary';
import { formatBytes } from '../features/datastudio/store/formatBytes';
import DashboardBar from '../features/datastudio/store/DashboardBar';
import TextAnalysis from '../features/datastudio/text/TextAnalysis';
import AutoBrief from '../features/datastudio/intent/AutoBrief';
import { useDatasetLibrary } from '../features/datastudio/store/useDashboards';
import { exportTilePng } from '../features/datastudio/store/exporters';

const ACCEPT = '.xlsx,.xlsm,.csv';

const TYPE_OPTIONS = ['numeric', 'categorical', 'multi', 'boolean', 'date', 'datetime', 'text', 'identifier'];
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
      <DatasetLibrary />
    </>
  );
}

/**
 * The storage-full dialog (spec §11).
 *
 * Quota exhaustion is the one storage failure a user can actually fix,
 * so it gets a real screen listing their datasets biggest-first with
 * delete buttons. The default browser behaviour is a silent failure
 * that is miserable to diagnose.
 */
function StorageFullDialog() {
  const { storageFull, dismissStorageFull } = useDataStudio();
  const { bySize, remove } = useDatasetLibrary();
  const [sized, setSized] = useState([]);

  useEffect(() => {
    if (!storageFull) return undefined;
    let cancelled = false;
    bySize().then((list) => { if (!cancelled) setSized(list); });
    return () => { cancelled = true; };
  }, [storageFull, bySize]);

  if (!storageFull) return null;

  return (
    <div className="ds-quota" role="alertdialog" aria-label="Browser storage is full">
      <Card className="ds-quota-card">
        <h3 className="ds-plan-heading">This browser has run out of storage</h3>
        <p className="ds-summary">
          Nothing was lost — the save just could not complete. Delete a dataset you no
          longer need and try again.
        </p>
        <ul className="ds-library-list">
          {sized.map((d) => (
            <li key={d.id}>
              <span className="ds-library-open">
                <span className="ds-library-name">{d.name}</span>
                <span className="ds-library-meta">{formatBytes(d.bytes)}</span>
              </span>
              <button
                type="button"
                className="ds-step-remove"
                aria-label={`Delete ${d.name} to free space`}
                onClick={() => { remove(d.id); setSized((c) => c.filter((x) => x.id !== d.id)); }}
              >
                <Trash2 size={14} />
              </button>
            </li>
          ))}
        </ul>
        <Button size="sm" variant="secondary" onClick={dismissStorageFull}>Close</Button>
      </Card>
    </div>
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
        <div className="ds-progress-bar" style={{ '--ds-progress': pct / 100 }} />
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

function CanvasStage() {
  const {
    tiles, dataset, editingTileId, addTile, setStage, reset, fileName, saveCurrentDataset,
    textColumns, analysing,
  } = useDataStudio();

  // Live ECharts instances, by tile id. Export needs a real chart object
  // to call getDataURL on, and nothing else exposes one.
  const chartsRef = useRef(new Map());
  const handleChartInit = useCallback((tileId, chart) => {
    chartsRef.current.set(tileId, chart);
  }, []);

  const handleExport = useCallback((tile) => {
    exportTilePng(chartsRef.current.get(tile.id), tile.title);
  }, []);

  if (!dataset) return <EmptyState>Nothing imported yet.</EmptyState>;

  return (
    <>
      <div className="ds-toolbar">
        <span className="ds-summary">
          {`${dataset.rowCount.toLocaleString()} rows · ${dataset.columns.length} columns`}
        </span>
        <span className="ds-toolbar-spacer" />
        <Button
          variant="secondary"
          size="sm"
          icon={Plus}
          onClick={() => addTile({
            title: 'New chart',
            chart: 'bar',
            encoding: { x: null, y: [{ column: null, agg: 'count' }], series: null },
            sort: { by: 'y', dir: 'desc' },
            limit: 10,
            size: 'M',
            stacked: false,
            respondsToFilters: true,
          })}
        >
          Add a chart
        </Button>
        {/* Only offered when the sheet actually holds written answers, so
            it never appears on a sheet of numbers. */}
        {textColumns.length > 0 && (
          <Button variant="secondary" size="sm" onClick={() => setStage('text')}>
            {analysing ? 'Text analysis (reading…)' : 'Text analysis'}
          </Button>
        )}
        <Button variant="secondary" size="sm" onClick={() => saveCurrentDataset(fileName)}>
          Save this data
        </Button>
        <Button variant="secondary" size="sm" onClick={() => setStage('cleaning')}>
          Back to cleaning
        </Button>
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reset}>
          Start over
        </Button>
      </div>

      <AutoBrief />
      <FilterBar />
      <DashboardBar />

      {tiles.length === 0 ? (
        <Card>
          <EmptyState>
            No charts yet. Use “Add a chart” to build one from your columns.
          </EmptyState>
        </Card>
      ) : (
        <div className={editingTileId ? 'ds-canvas-with-editor' : undefined}>
          <CanvasGrid onExport={handleExport} onChartInit={handleChartInit} />
          {editingTileId && <TileEditor />}
        </div>
      )}
    </>
  );
}

function DataStudioBody() {
  const { stage } = useDataStudio();
  if (stage === 'parsing') return <ParsingStage />;
  if (stage === 'idle') return <DropStage />;
  if (stage === 'cleaning') return <CleanReview />;
  if (stage === 'text') return <TextAnalysis />;
  if (stage === 'canvas') return <CanvasStage />;
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
        <StorageFullDialog />
      </AppShell>
    </DataStudioProvider>
  );
}
