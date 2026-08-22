import { useMemo } from 'react';
import Button from '../../../components/ui/Button';
import { X } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import { CHART_TYPES, chartTypeById } from './chartSpecs.js';

const AGGREGATIONS = ['sum', 'avg', 'count', 'countDistinct', 'min', 'max', 'median'];

/**
 * Edits one tile, previewing live.
 *
 * Columns are offered BY ROLE (spec §7.1): measures on Y, dimensions and
 * temporal columns on X. An identifier is never offered as a measure --
 * summing employee IDs produces a number, which is worse than producing
 * an error, because it looks like an answer.
 */
export default function TileEditor() {
  const {
    tiles, dataset, profile, editingTileId, setEditingTile, updateTile,
  } = useDataStudio();

  const tile = tiles.find((t) => t.id === editingTileId) ?? null;

  const columns = useMemo(() => profile?.columns ?? [], [profile]);
  const measures = useMemo(() => columns.filter((c) => c.role === 'measure'), [columns]);
  const axisColumns = useMemo(
    () => columns.filter((c) => c.role === 'dimension' || c.role === 'temporal'),
    [columns],
  );

  // Only chart types this dataset could actually satisfy. Offering
  // "Scatter" to someone with one measure just sets them up to pick it
  // and be told no.
  const availableTypes = useMemo(
    () => CHART_TYPES.filter((t) => measures.length >= t.needs.y || t.needs.y === 0),
    [measures.length],
  );

  if (!tile) return null;

  const type = chartTypeById(tile.chart);
  const isTemporalX = columns.find((c) => c.name === tile.encoding?.x?.column)?.role === 'temporal';

  const setEncoding = (patch) => updateTile(tile.id, {
    encoding: { ...tile.encoding, ...patch },
  });

  const setMeasure = (index, patch) => {
    const y = (tile.encoding.y ?? []).slice();
    y[index] = { ...y[index], ...patch };
    setEncoding({ y });
  };

  return (
    <aside className="ds-editor" aria-label={`Editing ${tile.title}`}>
      <header className="ds-editor-head">
        <h3>Edit tile</h3>
        <button type="button" aria-label="Close the tile editor" onClick={() => setEditingTile(null)}>
          <X size={15} />
        </button>
      </header>

      <label className="ds-field">
        <span>Title</span>
        <input
          className="ds-select"
          value={tile.title}
          onChange={(e) => updateTile(tile.id, { title: e.target.value })}
        />
      </label>

      <label className="ds-field">
        <span>Chart type</span>
        <select
          className="ds-select"
          value={tile.chart}
          onChange={(e) => updateTile(tile.id, { chart: e.target.value })}
        >
          {availableTypes.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
        </select>
      </label>

      {type?.needs.x > 0 && (
        <label className="ds-field">
          <span>X axis</span>
          <select
            className="ds-select"
            value={tile.encoding?.x?.column ?? ''}
            onChange={(e) => setEncoding({
              x: { ...tile.encoding.x, column: e.target.value },
            })}
          >
            <option value="">Pick a column</option>
            {axisColumns.map((c) => (
              <option key={c.name} value={c.name}>{`${c.name} (${c.role})`}</option>
            ))}
          </select>
        </label>
      )}

      {isTemporalX && (
        <label className="ds-field">
          <span>Group dates by</span>
          <select
            className="ds-select"
            value={tile.encoding?.x?.bin ?? 'day'}
            onChange={(e) => setEncoding({ x: { ...tile.encoding.x, bin: e.target.value } })}
          >
            <option value="day">Day</option>
            <option value="month">Month</option>
            <option value="quarter">Quarter</option>
            <option value="year">Year</option>
          </select>
        </label>
      )}

      {(tile.encoding?.y ?? []).map((measure, i) => (
        <div className="ds-editor-measure" key={`measure-${i}`}>
          <label className="ds-field">
            <span>{`Measure ${i + 1}`}</span>
            <select
              className="ds-select"
              value={measure.column ?? ''}
              onChange={(e) => setMeasure(i, { column: e.target.value || null })}
            >
              {/* An empty column with `count` is a row count, which is a
                  legitimate measure and the only one a sheet with no
                  numbers can offer. */}
              <option value="">Row count</option>
              {measures.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
            </select>
          </label>
          <label className="ds-field">
            <span>Aggregate</span>
            <select
              className="ds-select"
              value={measure.agg ?? 'sum'}
              onChange={(e) => setMeasure(i, { agg: e.target.value })}
            >
              {AGGREGATIONS.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </label>
        </div>
      ))}

      <label className="ds-field">
        <span>Split into series by</span>
        <select
          className="ds-select"
          value={tile.encoding?.series?.column ?? ''}
          onChange={(e) => setEncoding({
            series: e.target.value ? { column: e.target.value } : null,
          })}
        >
          <option value="">Nothing</option>
          {columns.filter((c) => c.role === 'dimension').map((c) => (
            <option key={c.name} value={c.name}>{c.name}</option>
          ))}
        </select>
      </label>

      <div className="ds-editor-row">
        <label className="ds-field">
          <span>Sort by</span>
          <select
            className="ds-select"
            value={tile.sort?.by ?? 'y'}
            onChange={(e) => updateTile(tile.id, { sort: { ...tile.sort, by: e.target.value } })}
          >
            <option value="y">Value</option>
            <option value="x">Category</option>
          </select>
        </label>
        <label className="ds-field">
          <span>Direction</span>
          <select
            className="ds-select"
            value={tile.sort?.dir ?? 'desc'}
            onChange={(e) => updateTile(tile.id, { sort: { ...tile.sort, dir: e.target.value } })}
          >
            <option value="desc">Descending</option>
            <option value="asc">Ascending</option>
          </select>
        </label>
        <label className="ds-field">
          <span>Show at most</span>
          <input
            className="ds-select"
            type="number"
            min="1"
            max="500"
            value={tile.limit ?? 10}
            onChange={(e) => updateTile(tile.id, { limit: Number(e.target.value) || 10 })}
          />
        </label>
      </div>

      <label className="ds-check">
        <input
          type="checkbox"
          checked={Boolean(tile.stacked)}
          onChange={(e) => updateTile(tile.id, { stacked: e.target.checked })}
        />
        <span>Stack the series</span>
      </label>

      <label className="ds-check">
        <input
          type="checkbox"
          checked={tile.respondsToFilters !== false}
          onChange={(e) => updateTile(tile.id, { respondsToFilters: e.target.checked })}
        />
        <span>
          Respond to filters
          <em>Turn this off to keep showing the unfiltered total for comparison.</em>
        </span>
      </label>

      {!dataset && <p className="ds-editor-note">No dataset loaded yet.</p>}

      <Button size="sm" variant="secondary" onClick={() => setEditingTile(null)}>
        Done
      </Button>
    </aside>
  );
}
