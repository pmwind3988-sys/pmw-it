import { useMemo, useRef } from 'react';
import { aggregate } from '../engine/aggregate.js';
import { toEChartsOption, validateTileSpec } from './chartSpecs.js';
import EChart from './EChart.jsx';

function formatValue(n, isPercent) {
  if (!Number.isFinite(n)) return '—';
  if (isPercent) return `${(n * 100).toFixed(1)}%`;
  if (Math.abs(n) >= 1000) return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
  return String(Math.round(n * 100) / 100);
}

/**
 * One tile: aggregate, render, emit clicks.
 *
 * A tile is a pure function of `(dataset, mask, tile)`. That is what
 * makes cross-filtering cheap -- the mask changes, this recomputes, and
 * nothing else in the tree has to know.
 */
export default function ChartTile({ tile, dataset, mask, onSelect, onChartInit }) {
  const chartRef = useRef(null);

  const validity = useMemo(() => validateTileSpec(tile, dataset), [tile, dataset]);

  const result = useMemo(
    () => (validity.ok ? aggregate(dataset, mask, tile) : null),
    [validity.ok, dataset, mask, tile],
  );

  const option = useMemo(
    () => (result ? toEChartsOption(tile.chart, result, tile) : null),
    [result, tile],
  );

  // Memoised, or `EChart`'s binding effect tears down and re-attaches on
  // every render.
  const onEvents = useMemo(() => ({
    click: (params) => {
      if (!onSelect) return;
      // The clicked category is the filter value. For a pie the name is
      // on the datum; for cartesian charts it is the axis label.
      const value = params?.name;
      if (value === undefined || value === null) return;
      onSelect({
        tileId: tile.id,
        column: tile.encoding?.x?.column,
        value,
      });
    },
  }), [onSelect, tile.id, tile.encoding]);

  const handleInit = useMemo(() => (chart) => {
    chartRef.current = chart;
    onChartInit?.(tile.id, chart);
  }, [onChartInit, tile.id]);

  if (!validity.ok) {
    return (
      <div className="ds-tile-body ds-tile-broken" role="status">
        <p>{validity.reason}</p>
        <p className="ds-tile-broken-hint">
          Edit this tile to point it at a column that exists, or remove it.
        </p>
      </div>
    );
  }

  if (result.categories.length === 0 && tile.chart !== 'kpi') {
    return (
      <div className="ds-tile-body ds-tile-empty" role="status">
        No rows match the current filters.
      </div>
    );
  }

  if (option.kind === 'kpi') {
    return (
      <div className="ds-tile-body ds-kpi">
        <span className="ds-kpi-value">{formatValue(option.value, tile.isPercent)}</span>
        <span className="ds-kpi-label">{option.label}</span>
      </div>
    );
  }

  if (option.kind === 'table') {
    return (
      <div className="ds-tile-body ds-tile-table-scroll">
        <table className="ds-table">
          <thead>
            <tr>
              {option.headers.map((h) => (
                <th key={h} className={h === option.headers[0] ? undefined : 'ds-num'}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {option.rows.map((row) => (
              <tr key={String(row[0])}>
                {row.map((cell, i) => (
                  <td key={`${row[0]}-${i}`} className={i === 0 ? undefined : 'ds-num'}>
                    {i === 0 ? cell : formatValue(cell, tile.isPercent)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  return (
    <EChart
      className="ds-tile-body ds-tile-chart"
      option={option}
      onEvents={onEvents}
      onInit={handleInit}
    />
  );
}
