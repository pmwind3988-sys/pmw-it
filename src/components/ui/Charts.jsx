import { Card, EmptyState } from './Surfaces';

/**
 * The two charts this app draws, lifted out of DashboardPage so the device
 * dashboard renders the same shapes rather than a second set that drifts.
 *
 * No chart library: these are CSS bars, and adding one for eight bar charts
 * would be the largest dependency in the app.
 */

/** Horizontal bars. Rows carrying an `onSelect` become the way into their slice. */
export function BarChart({ title, blurb, rows, onSelect, emptyText }) {
  const max = Math.max(1, ...rows.map((r) => r.value));
  return (
    <Card className="chart-card">
      <div className="chart-head">
        <h3>{title}</h3>
        <p>{blurb}</p>
      </div>
      {rows.length === 0 ? (
        <EmptyState>{emptyText}</EmptyState>
      ) : (
        rows.map((row) => {
          const bar = (
            <>
              <span className="bar-label" title={row.label}>
                {row.label}
              </span>
              <span className="bar-track">
                <span
                  className="bar-fill"
                  style={{ width: `${Math.round((row.value / max) * 100)}%`, background: row.color }}
                />
              </span>
              <span className="bar-value">{row.value}</span>
            </>
          );
          if (!onSelect) {
            return (
              <div className="bar-row" key={row.label}>
                {bar}
              </div>
            );
          }
          return (
            <button
              type="button"
              className="bar-row bar-row-btn"
              key={row.label}
              onClick={() => onSelect(row)}
              aria-label={`${row.label}: ${row.value}. Show these requests.`}
            >
              {bar}
            </button>
          );
        })
      )}
    </Card>
  );
}

/** Vertical bars — the only chart here where the x axis is time. */
export function ColumnChart({ title, blurb, columns }) {
  const max = Math.max(1, ...columns.map((c) => c.value));
  return (
    <Card className="chart-card">
      <div className="chart-head">
        <h3>{title}</h3>
        <p>{blurb}</p>
      </div>
      <div className="column-chart">
        {columns.map((column) => (
          <div className="column" key={column.label}>
            <span className="column-count">{column.value}</span>
            <span
              className="column-bar"
              style={{ height: `${Math.max(3, Math.round((column.value / max) * 100))}%` }}
            />
            <span className="column-label">{column.label}</span>
          </div>
        ))}
      </div>
    </Card>
  );
}
