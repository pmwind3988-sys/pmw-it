import Button from '../../../components/ui/Button';
import { Check, AlertTriangle } from '../../../components/ui/Icons';

/**
 * What each phase is doing, in words. `counted` phases get a real bar; the
 * others get an indeterminate one, because a bar that cannot move is worse
 * than one that admits it does not know.
 */
const PHASES = {
  starting: { text: () => 'Getting ready…', counted: false },
  provisioning: {
    text: (done, total) => (total
      ? `Preparing the SharePoint columns — ${done} of ${total}…`
      : 'Preparing the SharePoint columns…'),
    counted: true,
  },
  reading: { text: () => 'Reading the current register…', counted: false },
  writing: { text: (done, total) => `Saving ${done} of ${total}…`, counted: true },
  logging: { text: (done, total) => `Recording ${done} of ${total} changes…`, counted: true },
};

export default function SaveProgress({ state, onRetry, onDone }) {
  const {
    phase, done, total, results, error, changeCount, unchanged,
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
    const step = PHASES[phase] ?? PHASES.starting;
    const percent = step.counted && total ? Math.round((done / total) * 100) : 0;

    return (
      <div className="sp-status">
        <div
          className={`sp-bar${step.counted ? '' : ' sp-bar-waiting'}`}
          role="progressbar"
          aria-valuenow={step.counted ? done : undefined}
          aria-valuemin={step.counted ? 0 : undefined}
          aria-valuemax={step.counted ? total : undefined}
          aria-label={step.text(done, total)}
        >
          <span style={step.counted ? { width: `${percent}%` } : undefined} />
        </div>
        <p className="sp-detail">{step.text(done, total)}</p>
        {phase === 'provisioning' && (
          <p className="sp-detail sp-aside">
            First save only — the columns are created once, then reused.
          </p>
        )}
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
