import { formatScalar } from '../formatValue';
import { splitEntries, isMultiValue } from '../multiValue';

/**
 * A field that holds several things, drawn as one chip per thing rather than
 * as a run-on line. `limit` collapses the tail into a "+3 more" chip, which is
 * what the register wants; the device page passes no limit and shows them all.
 *
 * Within a chip, the parts of an entry (`Kingston | DDR4 | 3200 MHz`) stay
 * together, with everything after the first shown as its detail. Anything that
 * is not a multi-value field falls through to the plain rendering, so a caller
 * can hand every field to this component.
 */
export default function ValueCell({ value, fieldKey, kind, limit = 0 }) {
  if (!isMultiValue(fieldKey)) return formatScalar(value, kind);

  const entries = splitEntries(value, fieldKey);
  if (entries.length === 0) return '—';

  const shown = limit > 0 ? entries.slice(0, limit) : entries;
  const hidden = entries.length - shown.length;

  return (
    <span className="mv-list">
      {shown.map((entry, index) => (
        // Entries can repeat (two identical sticks of RAM), so the position is
        // the only stable key here.
        <span className="mv-item" key={`${entry.text}-${index}`} title={entry.text}>
          <span className="mv-main">{entry.parts[0] ?? entry.text}</span>
          {entry.parts.length > 1 && (
            <span className="mv-detail">{entry.parts.slice(1).join(' · ')}</span>
          )}
        </span>
      ))}
      {hidden > 0 && <span className="mv-more">+{hidden} more</span>}
    </span>
  );
}
