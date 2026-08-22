import { useState } from 'react';
import { Card, EmptyState, ErrorBanner } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { Trash2 } from '../../../components/ui/Icons';
import { formatMYT } from '../time/malaysiaTime.js';
import { useDataStudio } from '../useDataStudio';
import { useDatasetLibrary } from './useDashboards.js';
import { formatBytes } from './formatBytes.js';

/**
 * The saved-dataset library, under the drop zone on the idle screen.
 *
 * Everything here lives in this browser only, which the empty state says
 * out loud -- someone who cannot find yesterday's import on a different
 * machine deserves to know why rather than to conclude the app lost it.
 */
export default function DatasetLibrary() {
  const { openSavedDataset } = useDataStudio();
  const { datasets, estimate, loading, error, remove } = useDatasetLibrary();
  const [confirming, setConfirming] = useState(null);

  if (loading) return null;

  return (
    <Card className="ds-library">
      <div className="ds-library-head">
        <h3 className="ds-plan-heading">Saved in this browser</h3>
        {estimate.ratio !== null && (
          <span className="ds-summary">
            {`${formatBytes(estimate.usage)} of ${formatBytes(estimate.quota)} used`}
          </span>
        )}
      </div>

      {estimate.ratio !== null && (
        <div
          className="ds-progress-track"
          role="progressbar"
          aria-valuenow={Math.round(estimate.ratio * 100)}
          aria-valuemin={0}
          aria-valuemax={100}
          aria-label="Browser storage used"
        >
          <div
            className={`ds-progress-bar${estimate.ratio > 0.9 ? ' ds-progress-bar-full' : ''}`}
            style={{ width: `${Math.min(100, Math.round(estimate.ratio * 100))}%` }}
          />
        </div>
      )}

      {error && <ErrorBanner message={error} />}

      {datasets.length === 0 ? (
        <EmptyState>
          Nothing saved yet. Imports are kept in this browser only — they are never
          uploaded, and they will not follow you to another machine.
        </EmptyState>
      ) : (
        <ul className="ds-library-list">
          {datasets.map((d) => (
            <li key={d.id}>
              <button
                type="button"
                className="ds-library-open"
                onClick={() => openSavedDataset(d.id)}
              >
                <span className="ds-library-name">{d.name}</span>
                <span className="ds-library-meta">
                  {`${(d.rowCount ?? 0).toLocaleString()} rows · ${d.sourceFileName ?? 'unknown file'}`}
                  {d.importedAt ? ` · ${formatMYT(d.importedAt, 'date')}` : ''}
                </span>
              </button>

              {confirming === d.id ? (
                <span className="ds-library-confirm">
                  <span>Delete for good?</span>
                  <Button size="sm" variant="secondary" onClick={() => setConfirming(null)}>
                    Keep
                  </Button>
                  <Button
                    size="sm"
                    onClick={() => { remove(d.id); setConfirming(null); }}
                  >
                    Delete
                  </Button>
                </span>
              ) : (
                <button
                  type="button"
                  className="ds-step-remove"
                  aria-label={`Delete the saved dataset ${d.name}`}
                  onClick={() => setConfirming(d.id)}
                >
                  <Trash2 size={14} />
                </button>
              )}
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
}
