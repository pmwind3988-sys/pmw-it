import { serialList } from '../units';

/**
 * Every serial on a bulk line at once, newest item first.
 *
 * The pager answers "tell me about item 7"; this answers "which serials do we
 * have", which the pager could only do one press at a time. A row opens its
 * item in the pager below, so the list is a way in, not a second editor.
 */
export default function SerialList({ units, at, onPick }) {
  const { entries, missing } = serialList(units);
  if (!entries.length) return null;

  return (
    <div className="as-serials">
      <div className="as-serials-head">
        <strong>Serial numbers</strong>
        <span className="as-sub">
          {entries.length} recorded, newest first
          {missing > 0 && ` · ${missing} without one`}
        </span>
      </div>
      <ol className="as-serials-list">
        {entries.map((entry) => (
          <li key={entry.index}>
            <button
              type="button"
              className={`as-serials-row${entry.index === at ? ' is-at' : ''}`}
              onClick={() => onPick(entry.index)}
              aria-current={entry.index === at ? 'true' : undefined}
            >
              <span className="as-serials-no">#{entry.index + 1}</span>
              <span className="as-serials-code">{entry.serial || entry.label}</span>
              <span className="as-serials-meta">
                {[!entry.serial && 'label only', entry.condition, entry.status]
                  .filter(Boolean).join(' · ')}
              </span>
            </button>
          </li>
        ))}
      </ol>
    </div>
  );
}
