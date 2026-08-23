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
 *
 * `tone` colours the whole value and `entryTone` colours one chip at a time --
 * red for what needs attention, green for what does not. Both are optional and
 * both are off unless the caller asks, so the register stays plain.
 */
export default function ValueCell({ value, fieldKey, kind, limit = 0, tone = null, entryTone }) {
  const toneClass = (name) => (name ? ` vc-tone vc-tone-${name}` : '');

  if (!isMultiValue(fieldKey)) {
    const text = formatScalar(value, kind);
    return tone ? <span className={toneClass(tone).trim()}>{text}</span> : text;
  }

  const entries = splitEntries(value, fieldKey);
  if (entries.length === 0) return '—';

  const shown = limit > 0 ? entries.slice(0, limit) : entries;
  const hidden = entries.length - shown.length;

  return (
    <span className="mv-list">
      {shown.map((entry, index) => (
        // Entries can repeat (two identical sticks of RAM), so the position is
        // the only stable key here.
        <span
          className={`mv-item${toneClass(entryTone ? entryTone(entry.text) : tone)}`}
          key={`${entry.text}-${index}`}
          title={entry.text}
        >
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
