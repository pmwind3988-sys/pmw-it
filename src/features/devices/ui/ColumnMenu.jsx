import { useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown } from '../../../components/ui/Icons';
import { ALL_COLUMNS, COLUMN_SETS, COLUMN_SET_IDS, setMatching } from '../columnSets';

const CHOICES = ALL_COLUMNS.filter((column) => column.key !== 'computerName');

/**
 * The "Columns" button and its menu: start from one of the four sets, then
 * tick or untick any column. A set is remembered per browser; a hand-picked
 * list is not, so a refresh puts the list back on the last set chosen.
 *
 * Escape and a press outside close it and hand the focus back, as + Filter does.
 */
export default function ColumnMenu({ keys, onChooseSet, onToggle }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const wrap = useRef(null);
  const trigger = useRef(null);
  const search = useRef(null);

  useEffect(() => {
    if (!open) return undefined;

    const onPointer = (event) => {
      if (wrap.current && !wrap.current.contains(event.target)) setOpen(false);
    };
    const onKey = (event) => {
      if (event.key !== 'Escape') return;
      setOpen(false);
      trigger.current?.focus();
    };

    document.addEventListener('pointerdown', onPointer);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onPointer);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  useEffect(() => {
    if (open) search.current?.focus();
  }, [open]);

  const matched = setMatching(keys);
  const summary = matched ? COLUMN_SETS[matched].label.replace(/ \(\d+\)$/, '') : `Custom (${keys.length})`;

  const shown = useMemo(() => new Set(keys), [keys]);
  const needle = query.trim().toLowerCase();
  const visible = needle
    ? CHOICES.filter((column) => column.label.toLowerCase().includes(needle))
    : CHOICES;

  const toggleOpen = () => {
    if (open) setQuery('');
    setOpen(!open);
  };

  return (
    <div className="dtf dcm" ref={wrap}>
      <button
        ref={trigger}
        type="button"
        className="dcm-trigger"
        aria-expanded={open}
        aria-controls="dcm-menu"
        onClick={toggleOpen}
      >
        Columns: {summary}
        <ChevronDown size={14} />
      </button>

      {open && (
        <div className="dcm-menu" id="dcm-menu" role="region" aria-label="Choose columns">
          <div className="dcm-sets" role="group" aria-label="Start from a set">
            {COLUMN_SET_IDS.map((id) => (
              <button
                key={id}
                type="button"
                className="dcm-set"
                aria-pressed={matched === id}
                onClick={() => onChooseSet(id)}
              >
                {COLUMN_SETS[id].label}
              </button>
            ))}
          </div>

          <input
            ref={search}
            type="search"
            className="dcm-search"
            placeholder="Find a column"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            aria-label="Find a column"
          />

          <div className="dcm-list">
            {!needle && (
              <label className="dcm-item dcm-fixed">
                <input type="checkbox" checked disabled />
                Computer <span className="dcm-note">always shown</span>
              </label>
            )}
            {visible.map((column) => (
              <label key={column.key} className="dcm-item">
                <input
                  type="checkbox"
                  checked={shown.has(column.key)}
                  onChange={() => onToggle(column.key)}
                />
                {column.label}
              </label>
            ))}
            {visible.length === 0 && <p className="dcm-none">No column called that.</p>}
          </div>

          <p className="dcm-foot">
            {shown.size} of {CHOICES.length} shown. Your own picks reset when the page is refreshed.
          </p>
        </div>
      )}
    </div>
  );
}
