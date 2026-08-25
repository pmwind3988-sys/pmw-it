import { useCallback, useRef, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import { Inbox, RefreshCw } from '../components/ui/Icons';
import { SemanticProvider } from '../features/semantic/SemanticContext';
import { useSemantic } from '../features/semantic/useSemantic';
import CanvasGrid from '../features/semantic/canvas/CanvasGrid';
import FilterBar from '../features/semantic/canvas/FilterBar';
import ResponsePanel from '../features/semantic/canvas/ResponsePanel';
import TextAnalysis from '../features/semantic/text/TextAnalysis';
import AutoBrief from '../features/semantic/intent/AutoBrief';
import { exportTilePng } from '../features/semantic/export/exporters';

const ACCEPT = '.xlsx,.xlsm,.csv';

function DropStage() {
  const { importFile, error, fileName } = useSemantic();
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
      <Card className="sa-drop-card">
        <div
          className={`sa-drop${dragging ? ' sa-drop-active' : ''}`}
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
          <Inbox size={28} className="sa-drop-icon" />
          <p className="sa-drop-title">Drop a Forms export here</p>
          <p className="sa-drop-hint">
            The <code>.xlsx</code> or <code>.csv</code> you download from Microsoft Forms.
            The written answers are read, sorted into categories and charted for you —
            all of it in this browser. Nothing is uploaded and nothing is saved:
            close the tab and it is gone.
          </p>
          <Button variant="secondary" onClick={() => inputRef.current?.click()}>
            Choose a file
          </Button>
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPT}
            className="sa-drop-input"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) importFile(file);
              e.target.value = '';
            }}
          />
          {fileName && !error && <p className="sa-drop-hint">Last file: {fileName}</p>}
        </div>
      </Card>
    </>
  );
}

function ParsingStage() {
  const { progress, fileName } = useSemantic();
  const pct = Math.max(0, Math.min(100, progress.pct ?? 0));
  return (
    <Card className="sa-progress-card">
      <p className="sa-progress-file">{fileName}</p>
      <div
        className="sa-progress-track"
        role="progressbar"
        aria-valuenow={pct}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-label={progress.stage || 'Reading'}
      >
        <div className="sa-progress-bar" style={{ '--sa-progress': pct / 100 }} />
      </div>
      <p className="sa-progress-stage">{progress.stage || 'Working...'}</p>
    </Card>
  );
}

/**
 * The charts, and the responses behind them.
 *
 * There is no chart builder and no cleaning checklist: every chart on
 * this screen was chosen from the sheet, and the ones about categories,
 * severity and themes were chosen from what the model read out of the
 * written answers. The only editing on offer is the destructive kind —
 * resize a chart, remove one — because everything else the user might
 * want to change is a question about the ANALYSIS, and that is what the
 * other tab is.
 */
function DashboardStage() {
  const {
    tiles, dataset, sheets, activeSheet, selectSheet, reset, setStage,
    headerIndex, headerCandidates, setHeaderIndex,
    textColumns, analysing, analysis,
  } = useSemantic();

  // Live ECharts instances, by tile id. Export needs a real chart
  // object to call getDataURL on, and nothing else exposes one.
  const chartsRef = useRef(new Map());
  const handleChartInit = useCallback((tileId, chart) => {
    chartsRef.current.set(tileId, chart);
  }, []);

  const handleExport = useCallback((tile) => {
    exportTilePng(chartsRef.current.get(tile.id), tile.title);
  }, []);

  if (!dataset) return <EmptyState>Nothing read yet.</EmptyState>;

  return (
    <>
      <div className="sa-toolbar">
        {/* A Forms export is one sheet, but a workbook somebody has
            added a working sheet to is not, and the wrong one reads as
            an empty screen. */}
        {sheets.length > 1 && (
          <label className="sa-field">
            <span>Sheet</span>
            <select
              className="sa-select"
              value={activeSheet}
              onChange={(e) => selectSheet(e.target.value)}
            >
              {sheets.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
        )}
        {/* Detection gets this right on a Forms export every time. It
            is here for the workbook somebody has typed a title across
            the top of, where the questions are on row 3 and every
            chart would otherwise be drawn from the wrong labels. */}
        {headerCandidates.length > 1 && (
          <details className="sa-headerpick">
            <summary>Questions are on row {headerIndex + 1}</summary>
            <label className="sa-field">
              <span className="sa-sr-only">Which row holds the question titles</span>
              <select
                className="sa-select"
                value={headerIndex}
                onChange={(e) => setHeaderIndex(Number(e.target.value))}
              >
                {headerCandidates.map((row, i) => (
                  <option key={row.join('|') || i} value={i}>
                    {`Row ${i + 1} — ${row.filter(Boolean).slice(0, 4).join(', ') || '(blank)'}`}
                  </option>
                ))}
              </select>
            </label>
          </details>
        )}
        <span className="sa-summary">
          {`${dataset.rowCount.toLocaleString()} responses · ${dataset.columns.length} fields`}
        </span>
        <span className="sa-toolbar-spacer" />
        {textColumns.length > 0 && (
          <Button
            variant="secondary"
            size="sm"
            onClick={() => setStage('text')}
            disabled={!analysis && !analysing}
          >
            {analysing ? 'Reading the answers…' : 'Categories and issues'}
          </Button>
        )}
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={reset}>
          Start over
        </Button>
      </div>

      <AutoBrief />
      <FilterBar />

      {tiles.length === 0 ? (
        <Card>
          <EmptyState>
            Nothing in this sheet could be charted. The responses are still below.
          </EmptyState>
        </Card>
      ) : (
        <CanvasGrid onExport={handleExport} onChartInit={handleChartInit} />
      )}

      {/* Below the charts, not beside them: the responses are the
          answer to "who is in that bar", and that question only comes
          up once the bar has been tapped. */}
      <ResponsePanel />
    </>
  );
}

function SemanticBody() {
  const { stage } = useSemantic();
  if (stage === 'parsing') return <ParsingStage />;
  if (stage === 'text') return <TextAnalysis />;
  if (stage === 'dashboard') return <DashboardStage />;
  return <DropStage />;
}

export default function SemanticAnalysisPage() {
  return (
    <SemanticProvider>
      <AppShell
        title="Semantic Analysis"
        subtitle="Drop a Microsoft Forms export and read what people actually wrote."
      >
        <SemanticBody />
      </AppShell>
    </SemanticProvider>
  );
}
