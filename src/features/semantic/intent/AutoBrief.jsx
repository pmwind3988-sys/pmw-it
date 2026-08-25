import { useMemo, useState } from 'react';
import { Card } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { BarChart3, ChevronRight, X } from '../../../components/ui/Icons';
import { useSemantic } from '../useSemantic';
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
 * Three decisions are taken on the user's behalf -- what the sheet is
 * about, which columns to park as bookkeeping, and which written-answer
 * column to read -- and this card is the whole of the user's defence
 * against any of them being wrong. Every decision is named, and the
 * reversal is one click rather than a re-import.
 *
 * It lives in `intent/` rather than `canvas/` because it renders the
 * plan, not the charts. Nothing among the charts knows it exists.
 */
export default function AutoBrief() {
  const {
    brief, profile, grid, analysis, analysing, textProgress, textError,
    showHiddenColumns, hideAdminColumns, dismissBrief, setStage, startAnalysis,
  } = useSemantic();
  const [expanded, setExpanded] = useState(false);
  const summary = useMemo(() => (analysis ? summarize(analysis) : ''), [analysis]);

  if (!brief || brief.dismissed) return null;

  const hidden = brief.hidden ?? [];
  const shown = brief.hiddenShown === true;
  const charted = (profile?.columns ?? []).filter((c) => c.role !== 'ignored').length;
  const rows = grid?.rows?.length ?? 0;

  return (
    <Card className="sa-brief">
      <div className="sa-brief-head">
        <BarChart3 size={16} className="sa-brief-icon" />
        <div className="sa-brief-titles">
          <h3 className="sa-brief-title">{brief.intent.title || 'Your form'}</h3>
          <p className="sa-brief-line">
            {`Read as ${brief.intent.label}. `}
            {`${rows.toLocaleString()} responses · ${plural(charted, 'field', 'fields')} charted`}
            {hidden.length > 0 && !shown ? ` · ${hidden.length} set aside` : ''}
          </p>
        </div>
        <button
          type="button"
          className="sa-brief-dismiss"
          onClick={dismissBrief}
          aria-label="Hide this summary"
        >
          <X size={14} />
        </button>
      </div>

      {hidden.length > 0 && (
        <div className="sa-brief-hidden">
          <button
            type="button"
            className="sa-brief-toggle"
            onClick={() => setExpanded((open) => !open)}
            aria-expanded={expanded}
          >
            <ChevronRight
              size={13}
              className={`sa-brief-chevron${expanded ? ' sa-brief-chevron-open' : ''}`}
            />
            {shown
              ? `${plural(hidden.length, 'bookkeeping column is', 'bookkeeping columns are')} being charted`
              : `${plural(hidden.length, 'bookkeeping column was', 'bookkeeping columns were')} left out`}
          </button>

          {expanded && (
            <>
              <ul className="sa-brief-list">
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
        <div className="sa-brief-text">
          {analysing && (
            <>
              <p className="sa-brief-line">
                {`Reading the answers to “${brief.analyseColumn}” — ${textProgress.stage || 'starting'}`}
              </p>
              <div className="sa-progress-track">
                <div
                  className="sa-progress-bar"
                  style={{ '--sa-progress': (textProgress.pct ?? 0) / 100 }}
                />
              </div>
            </>
          )}

          {!analysing && analysis && (
            <div className="sa-brief-row">
              <p className="sa-brief-line">
                {`${summary} — from the answers to “${analysis.columnName}”.`}
              </p>
              <Button variant="ghost" size="sm" onClick={() => setStage('text')}>
                Open the analysis
              </Button>
            </div>
          )}

          {/* Only reachable when the reading failed or was stopped:
              an import starts it by itself. */}
          {!analysing && !analysis && (
            <div className="sa-brief-row">
              <p className="sa-brief-line">
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
