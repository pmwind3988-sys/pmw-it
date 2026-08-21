import Button from '../../../components/ui/Button';
import { Check, AlertTriangle } from '../../../components/ui/Icons';

export default function SaveProgress({ state, onRetry, onDone }) {
  const {
    done, total, results, error, changeCount, unchanged,
  } = state;

  if (error) {
    return (
      <div className="sp-status">
        <AlertTriangle size={20} className="sp-icon-bad" />
        <p className="sp-headline">Nothing was saved</p>
        <p className="sp-detail">{error}</p>
        <Button variant="secondary" onClick={() => onRetry(null)}>Try again</Button>
      </div>
    );
  }

  if (results === null) {
    const percent = total ? Math.round((done / total) * 100) : 0;
    return (
      <div className="sp-status">
        <div
          className="sp-bar"
          role="progressbar"
          aria-valuenow={done}
          aria-valuemin={0}
          aria-valuemax={total}
        >
          <span style={{ width: `${percent}%` }} />
        </div>
        <p className="sp-detail">Saving {done} of {total}…</p>
      </div>
    );
  }

  const failures = results.filter((row) => row.error);
  const saved = results.length - failures.length;
  const inserted = results.filter((row) => !row.error && row.action === 'insert').length;
  const updated = results.filter((row) => !row.error && row.action === 'update').length;

  return (
    <div className="sp-status">
      {failures.length === 0
        ? <Check size={20} className="sp-icon-good" />
        : <AlertTriangle size={20} className="sp-icon-bad" />}

      <p className="sp-headline">
        {saved === 0 && failures.length === 0 ? 'Nothing to save' : `${saved} saved`}
      </p>

      <p className="sp-detail">
        {inserted > 0 && `${inserted} added`}
        {inserted > 0 && updated > 0 && ' · '}
        {updated > 0 && `${updated} updated`}
        {changeCount > 0 && ` · ${changeCount} change${changeCount === 1 ? '' : 's'} logged`}
        {unchanged > 0 && ` · ${unchanged} already current`}
      </p>

      {failures.length > 0 && (
        <ul className="sp-failures">
          {failures.map((row) => (
            <li key={row.computerName}>
              <strong>{row.computerName}</strong> — {row.error}
            </li>
          ))}
        </ul>
      )}

      <div className="sp-actions">
        {failures.length > 0 && (
          <Button
            variant="secondary"
            onClick={() => onRetry(failures.map((row) => row.computerName))}
          >
            Retry the {failures.length} that failed
          </Button>
        )}
        <Button onClick={onDone}>Done</Button>
      </div>
    </div>
  );
}
