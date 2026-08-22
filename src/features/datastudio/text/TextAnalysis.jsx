import { useState } from 'react';
import { Card, EmptyState, ErrorBanner } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { RefreshCw, BarChart3 } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import BucketEditor from './BucketEditor';
import IssueTable from './IssueTable';
import ThemeList from './ThemeList';
import PriorityBoard from './PriorityBoard';
import { UNSORTED_ID } from './buckets';

const VIEWS = [
  ['buckets', 'Categories'],
  ['issues', 'Issues'],
  ['themes', 'Themes'],
  ['priority', 'Priority'],
];

/**
 * The Text Analysis stage.
 *
 * Two things on this screen are deliberate and easy to "tidy away":
 *
 *   * the progress bar names the stage it is in. Loading the model takes
 *     the better part of ten seconds the first time, and an unlabelled
 *     bar reads as a hang.
 *   * "everything landed in Unsorted" gets its own message with the fix
 *     in it. The alternative is a screen that looks broken but is in
 *     fact the model being honest about not recognising anything.
 */
export default function TextAnalysis() {
  const {
    textColumns, textColumnName, analysis, analysing, textProgress, textError,
    startAnalysis, applyAnalysisColumns, resetOverrides, setStage,
  } = useDataStudio();
  const [view, setView] = useState('buckets');

  if (textColumns.length === 0) {
    return (
      <EmptyState>
        Nothing in this sheet is long enough to read as written answers.
      </EmptyState>
    );
  }

  const unsorted = analysis?.buckets.find((b) => b.id === UNSORTED_ID)?.count ?? 0;
  const counted = analysis?.fragments.filter((f) => !f.noise) ?? [];
  const people = new Set(counted.map((f) => f.row)).size;
  const allUnsorted = counted.length > 0 && unsorted === counted.length;

  return (
    <>
      {textError && (
        <ErrorBanner message={textError} onRetry={() => startAnalysis(textColumnName)} />
      )}

      <div className="ds-toolbar">
        <label className="ds-field">
          <span>Read</span>
          <select
            className="ds-select"
            value={textColumnName}
            onChange={(e) => startAnalysis(e.target.value)}
          >
            <option value="">Choose a question…</option>
            {textColumns.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </label>

        <span className="ds-toolbar-spacer" />

        {analysis && (
          <span className="ds-summary">
            {`${counted.length} issues from ${people} people · `}
            {`${analysis.themes.length} themes · ${unsorted} unsorted`}
          </span>
        )}

        <Button variant="ghost" size="sm" onClick={resetOverrides} disabled={!analysis}>
          Reset my edits
        </Button>
        <Button
          variant="ghost"
          size="sm"
          icon={BarChart3}
          onClick={applyAnalysisColumns}
          disabled={!analysis}
        >
          Add to my charts
        </Button>
        <Button variant="ghost" size="sm" icon={RefreshCw} onClick={() => setStage('canvas')}>
          Back to charts
        </Button>
      </div>

      {!analysis && !analysing && (
        <Card>
          <EmptyState>
            <p>Read the written answers and sort them into categories.</p>
            <p className="ds-drop-hint">
              This runs on your machine. Nothing is uploaded. The first run loads the
              model once, which takes a few seconds.
            </p>
            <Button onClick={() => startAnalysis(textColumnName || textColumns[0].name)}>
              Analyse the answers
            </Button>
          </EmptyState>
        </Card>
      )}

      {analysing && (
        <Card>
          <p className="ds-summary">{textProgress.stage || 'Working'}</p>
          <div className="bar-track">
            <span
              className="bar-fill"
              style={{ transform: `scaleX(${(textProgress.pct ?? 0) / 100})` }}
            />
          </div>
        </Card>
      )}

      {analysis && (
        <>
          {allUnsorted && (
            <Card className="ds-text-notice">
              <p className="ds-summary">
                Nothing matched a category confidently. Lower the confidence setting on the
                Categories tab to accept weaker matches, or describe your categories in
                more detail — the descriptions are what answers are matched against.
              </p>
            </Card>
          )}

          <div className="ds-text-tabs" role="tablist">
            {VIEWS.map(([id, label]) => (
              <button
                key={id}
                type="button"
                role="tab"
                aria-selected={view === id}
                className={`ds-text-tab${view === id ? ' ds-text-tab-on' : ''}`}
                onClick={() => setView(id)}
              >
                {label}
              </button>
            ))}
          </div>

          {view === 'buckets' && <BucketEditor />}
          {view === 'issues' && <IssueTable />}
          {view === 'themes' && <ThemeList />}
          {view === 'priority' && <PriorityBoard />}
        </>
      )}
    </>
  );
}
