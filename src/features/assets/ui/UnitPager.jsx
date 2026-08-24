import { useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, Boxes, Check } from '../../../components/ui/Icons';
import {
  UNIT_FIELDS, setUnitField, isBlankUnit, filledCount,
} from '../units';

/**
 * A bulk row, one physical item at a time.
 *
 * The register says "2 tabs" because that is what was bought. This is where
 * the second tab gets its own serial number, its own sticker label and its own
 * "screen cracked" — without splitting the line into two rows and losing the
 * fact that it was one delivery of two identical things.
 *
 * It pages rather than scrolls. Five identical cards stacked down a phone
 * screen are impossible to tell apart, and the question being answered is
 * always about ONE of them; a swipe left and right is how that thing is held
 * and turned over in real life.
 */

/** Below this a drag is a scroll, not a swipe. */
const SWIPE_PX = 45;

export default function UnitPager({ units, onChange }) {
  const [at, setAt] = useState(0);
  const touch = useRef(null);

  const count = units.length;
  // Clamped during render rather than in an effect: lowering the quantity from
  // five to two while sitting on unit five must not paint a frame of nothing.
  const index = Math.min(at, count - 1);
  const unit = units[index];

  const go = (to) => setAt(Math.max(0, Math.min(count - 1, to)));

  const set = (field) => (event) => onChange(
    setUnitField(units, unit.index, field, event.target.value),
  );

  const onTouchStart = (event) => {
    touch.current = { x: event.touches[0].clientX, y: event.touches[0].clientY };
  };

  /**
   * A vertical drag is the page scrolling and must not turn the card. The
   * comparison against the vertical distance is what stops the pager from
   * stealing every scroll that starts on top of it.
   */
  const onTouchEnd = (event) => {
    const start = touch.current;
    touch.current = null;
    if (!start) return;

    const dx = event.changedTouches[0].clientX - start.x;
    const dy = event.changedTouches[0].clientY - start.y;
    if (Math.abs(dx) < SWIPE_PX || Math.abs(dx) < Math.abs(dy)) return;

    go(index + (dx < 0 ? 1 : -1));
  };

  const recorded = filledCount(units);

  return (
    <div
      className="as-units"
      onTouchStart={onTouchStart}
      onTouchEnd={onTouchEnd}
    >
      <div className="as-units-bar">
        <button
          type="button"
          className="as-iconbtn"
          onClick={() => go(index - 1)}
          disabled={index === 0}
          aria-label="Previous item"
        >
          <ChevronLeft size={15} />
        </button>

        <div className="as-units-which">
          <strong>Item {index + 1} of {count}</strong>
          <span className="as-sub">
            {recorded === 0
              ? 'Nothing recorded against any of them yet'
              : `${recorded} of ${count} filled in`}
          </span>
        </div>

        <button
          type="button"
          className="as-iconbtn"
          onClick={() => go(index + 1)}
          disabled={index === count - 1}
          aria-label="Next item"
        >
          <ChevronRight size={15} />
        </button>
      </div>

      {/* Dots, so twenty cables do not need twenty presses to reach the one
          that matters, and so a filled unit is findable at a glance. */}
      {count > 1 && (
        <div className="as-units-dots" role="tablist" aria-label="Items on this row">
          {units.map((entry, position) => (
            <button
              key={entry.index}
              type="button"
              role="tab"
              aria-selected={position === index}
              aria-label={`Item ${position + 1}`}
              className={[
                'as-units-dot',
                position === index ? 'is-at' : '',
                isBlankUnit(entry) ? '' : 'is-filled',
              ].filter(Boolean).join(' ')}
              onClick={() => go(position)}
            />
          ))}
        </div>
      )}

      <div className="as-form">
        {UNIT_FIELDS.map((field) => (
          <label className="as-field" key={field.key}>
            <span className="as-field-label">{field.label}</span>

            {field.options ? (
              <select value={unit[field.key]} onChange={set(field.key)}>
                {/* Empty is "nobody has said", not a value. The row has no
                    condition of its own to fall back to — it is a count of
                    things, and only a thing can be faulty. */}
                <option value="">— not recorded</option>
                {field.options.map((option) => (
                  <option key={option} value={option}>{option}</option>
                ))}
              </select>
            ) : field.multiline ? (
              <textarea rows={2} value={unit[field.key]} onChange={set(field.key)} />
            ) : (
              <input
                value={unit[field.key]}
                onChange={set(field.key)}
                placeholder={field.key === 'serialNumber' ? 'The serial on this one' : ''}
              />
            )}
          </label>
        ))}
      </div>

      <p className="as-units-foot">
        <Boxes size={13} />
        <span>
          Swipe, or use the arrows, to reach the other {count === 2 ? 'one' : 'ones'}.
          {' '}Changes to every item are saved together with Save changes.
        </span>
        {!isBlankUnit(unit) && <Check size={13} className="as-units-tick" />}
      </p>
    </div>
  );
}
