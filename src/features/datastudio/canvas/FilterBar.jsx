import { useMemo, useState } from 'react';
import Button from '../../../components/ui/Button';
import { X, Filter } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';

/**
 * What is currently narrowing the dashboard, and how to undo it.
 *
 * The cross-filter selection gets its own chip, distinct from the
 * global filters. Without it a user who clicked a bar and then scrolled
 * has a dashboard filtered by something with no visible cause -- and
 * the fix, clicking that same mark again, is off screen.
 */
export default function FilterBar() {
  const {
    profile, tiles, globalFilters, addFilter, removeFilter, clearFilters,
    selection, clearSelection,
  } = useDataStudio();

  const [pendingColumn, setPendingColumn] = useState('');

  const filterable = useMemo(
    () => (profile?.columns ?? []).filter(
      (c) => c.role === 'dimension' && c.topValues.length > 0,
    ),
    [profile],
  );

  const pending = filterable.find((c) => c.name === pendingColumn) ?? null;
  const sourceTile = selection
    ? tiles.find((t) => t.id === selection.sourceTileId)
    : null;

  const hasAnything = globalFilters.length > 0 || Boolean(selection);

  return (
    <div className="ds-filterbar">
      <Filter size={14} className="ds-filterbar-icon" />

      {selection && (
        <span className="ds-chip ds-chip-selection">
          <span>
            {`${selection.column}: ${selection.values.join(', ')}`}
            <em>{` from ${sourceTile?.title ?? 'a chart'}`}</em>
          </span>
          <button type="button" aria-label="Clear the chart selection" onClick={clearSelection}>
            <X size={12} />
          </button>
        </span>
      )}

      {globalFilters.map((filter) => (
        <span className="ds-chip" key={filter.column}>
          <span>{`${filter.column}: ${(filter.values ?? []).join(', ')}`}</span>
          <button
            type="button"
            aria-label={`Remove the ${filter.column} filter`}
            onClick={() => removeFilter(filter.column)}
          >
            <X size={12} />
          </button>
        </span>
      ))}

      <label className="ds-filterbar-add">
        <span className="ds-sr-only">Add a filter</span>
        <select
          className="ds-select"
          value={pendingColumn}
          onChange={(e) => setPendingColumn(e.target.value)}
        >
          <option value="">Add a filter…</option>
          {filterable.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
        </select>
      </label>

      {pending && (
        <label className="ds-filterbar-add">
          <span className="ds-sr-only">{`Value for ${pending.name}`}</span>
          <select
            className="ds-select"
            value=""
            onChange={(e) => {
              if (!e.target.value) return;
              addFilter({ column: pending.name, kind: 'in', values: [e.target.value] });
              setPendingColumn('');
            }}
          >
            <option value="">Pick a value…</option>
            {pending.topValues.map((v) => (
              <option key={v.value} value={v.value}>{`${v.value} (${v.count})`}</option>
            ))}
          </select>
        </label>
      )}

      {hasAnything && (
        <Button
          size="sm"
          variant="secondary"
          onClick={() => { clearFilters(); clearSelection(); }}
        >
          Clear all
        </Button>
      )}
    </div>
  );
}
