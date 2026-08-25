import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { fitByDepartment } from '../stats/deviceStats';

/**
 * Which department is in the most trouble, and what kind of trouble it is.
 *
 * One row per department, each row a bar split into the four levels. Reading
 * left to right, the redder the left-hand end, the sooner somebody has to go
 * and look. Every segment opens the machines it counted.
 */
const SEGMENTS = [
  { key: 'Critical', color: 'var(--it-danger)' },
  { key: 'Needs Attention', color: 'var(--it-accent)' },
  { key: 'Moderate', color: 'var(--it-brand)' },
  { key: 'Optimal', color: 'var(--it-good)' },
  { key: 'Unknown', color: 'var(--it-ink-soft)' },
];

export default function DepartmentHeatmap({ devices, onSelect }) {
  const rows = fitByDepartment(devices);

  return (
    <Card className="chart-card dv-heat">
      <div className="chart-head">
        <h3>Department risk</h3>
        <p>
          How each department&apos;s machines measure up to the work it does.
          Click a block to see the machines in it.
        </p>
      </div>

      {rows.length === 0 ? (
        <EmptyState>No devices imported yet.</EmptyState>
      ) : (
        <>
          <ul className="dv-heat-rows">
            {rows.map((row) => (
              <li key={row.department} className="dv-heat-row">
                <span className="dv-heat-name" title={`${row.department} — ${row.persona}`}>
                  {row.department}
                </span>
                <span className="dv-heat-bar">
                  {SEGMENTS.filter((segment) => row[segment.key] > 0).map((segment) => (
                    <button
                      type="button"
                      key={segment.key}
                      className="dv-heat-cell"
                      style={{
                        width: `${(row[segment.key] / row.total) * 100}%`,
                        background: segment.color,
                      }}
                      onClick={() => onSelect?.(row.department, segment.key)}
                      title={`${row.department}: ${row[segment.key]} ${segment.key}`}
                      aria-label={`${row.department}, ${segment.key}: ${row[segment.key]} machines. Show them.`}
                    >
                      {row[segment.key]}
                    </button>
                  ))}
                </span>
                <span className="dv-heat-index" title="Share of the department's machines that do not suit its work">
                  {row.riskIndex}%
                </span>
              </li>
            ))}
          </ul>

          <ul className="dv-legend">
            {SEGMENTS.map((segment) => (
              <li key={segment.key}>
                <span className="dv-legend-dot" style={{ background: segment.color }} />
                {segment.key}
              </li>
            ))}
          </ul>
        </>
      )}
    </Card>
  );
}
