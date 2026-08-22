import { useMemo, useState } from 'react';
import { Card } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { BarChart3, ChevronRight, X } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import { UNSORTED_ID } from '../text/buckets';

/**
 * The one-line result of reading a column.
 *
 * Categories rather than themes as the headline number. Themes are
 * discovered, so a survey where everybody raised a different problem
 * honestly has none -- and "0 themes" reads as a failed analysis when
 * in fact every answer was filed. The theme count is added only when
 * there is one to report.
 */
const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

function summarize(analysis) {
  const filed = new Set();
  for (const fragment of analysis.fragments) {
    if (fragment.bucketId && fragment.bucketId !== UNSORTED_ID) filed.add(fragment.bucketId);
  }

  const parts = [
    plural(analysis.fragments.length, 'issue', 'issues'),
    plural(filed.size, 'category', 'categories'),
  ];
  if (analysis.themes.length > 0) {
    parts.push(plural(analysis.themes.length, 'repeated theme', 'repeated themes'));
  }
  return parts.join(' · ');
}

/**
 * What the app did to the file before the user saw it.
 *
 * The autopilot takes four decisions on the user's behalf -- what the
 * sheet is about, which columns to park, which charts to draw, whether
 * to read the written answers -- and this card is the whole of the
 * user's defence against any of them being wrong. Every decision is
 * named, every one is reversible from here, and the reversal is one
 * click rather than a re-import.
 *
 * It lives in `intent/` rather than `canvas/` because it renders the
 * plan, not the charts. Nothing on the canvas knows it exists.
 */
export default function AutoBrief() {
  const {
    brief, profile, grid, analysis, analysing, textProgress, textError,
    showHiddenColumns, hideAdminColumns, dismissBrief, setStage, startAnalysis,
  } = useDataStudio();
  const [expanded, setExpanded] = useState(false);
  const summary = useMemo(() => (analysis ? summarize(analysis) : ''), [analysis]);

  if (!brief || brief.dismissed) return null;

  const hidden = brief.hidden ?? [];
  const shown = brief.hiddenShown === true;
  const charted = (profile?.columns ?? []).filter((c) => c.role !== 'ignored').length;
  const rows = grid?.rows?.length ?? 0;

  return (
    <Card className="ds-brief">
      <div className="ds-brief-head">
        <BarChart3 size={16} className="ds-brief-icon" />
        <div className="ds-brief-titles">
          <h3 className="ds-brief-title">{brief.intent.title || 'Your spreadsheet'}</h3>
          <p className="ds-brief-line">
            {`Read as ${brief.intent.label}. `}
            {`${rows.toLocaleString()} rows · ${plural(charted, 'column', 'columns')} charted`}
            {hidden.length > 0 && !shown ? ` · ${hidden.length} set aside` : ''}
          </p>
        </div>
        <Button variant="subtle" size="sm" onClick={() => setStage('profiled')}>
          Review every column
        </Button>
        <button
          type="button"
          className="ds-brief-dismiss"
          onClick={dismissBrief}
          aria-label="Hide this summary"
        >
          <X size={14} />
        </button>
      </div>

      {hidden.length > 0 && (
        <div className="ds-brief-hidden">
          <button
            type="button"
            className="ds-brief-toggle"
            onClick={() => setExpanded((open) => !open)}
            aria-expanded={expanded}
          >
            <ChevronRight
              size={13}
              className={`ds-brief-chevron${expanded ? ' ds-brief-chevron-open' : ''}`}
            />
            {shown
              ? `${plural(hidden.length, 'bookkeeping column is', 'bookkeeping columns are')} being charted`
              : `${plural(hidden.length, 'bookkeeping column was', 'bookkeeping columns were')} left out`}
          </button>

          {expanded && (
            <>
              <ul className="ds-brief-list">
                {hidden.map((column) => (
                  <li key={column.name}>
                    <strong>{column.name || '(no header)'}</strong>
                    <span> — {column.reason}</span>
                  </li>
                ))}
              </ul>
              <Button
                variant="ghost"
                size="sm"
                onClick={shown ? hideAdminColumns : showHiddenColumns}
              >
                {shown ? 'Leave them out again' : 'Chart them anyway'}
              </Button>
            </>
          )}
        </div>
      )}

      {brief.analyseColumn && (
        <div className="ds-brief-text">
          {analysing && (
            <>
              <p className="ds-brief-line">
                {`Reading the answers to “${brief.analyseColumn}” — ${textProgress.stage || 'starting'}`}
              </p>
              <div className="ds-progress-track">
                <div
                  className="ds-progress-bar"
                  style={{ '--ds-progress': (textProgress.pct ?? 0) / 100 }}
                />
              </div>
            </>
          )}

          {!analysing && analysis && (
            <div className="ds-brief-row">
              <p className="ds-brief-line">
                {`${summary} — from the answers to “${analysis.columnName}”.`}
              </p>
              <Button variant="ghost" size="sm" onClick={() => setStage('text')}>
                Open the analysis
              </Button>
            </div>
          )}

          {/* Offered rather than started: reading the writing pulls a
              23MB model the first time, and a sheet whose title is not
              about written answers has not asked for that. */}
          {!analysing && !analysis && (
            <div className="ds-brief-row">
              <p className="ds-brief-line">
                {textError
                  || `“${brief.analyseColumn}” holds written answers that can be grouped by meaning.`}
              </p>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => startAnalysis(brief.analyseColumn)}
              >
                {textError ? 'Try again' : 'Read them'}
              </Button>
            </div>
          )}
        </div>
      )}
    </Card>
  );
}
