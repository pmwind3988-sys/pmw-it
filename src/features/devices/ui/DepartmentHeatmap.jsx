import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { criticalPartsByDepartment } from '../stats/deviceStats';
import { PARTS } from '../standards/defaultStandard';

const MIN_DOT = 22;
const MAX_DOT = 40;

/**
 * Which department is in the most trouble, and on which part. Each non-zero
 * cell is a red dot whose size follows its count; pressing it opens those
 * machines. Colour is the Critical grade, never a fixed red.
 */
export default function DepartmentHeatmap({ devices, onSelect }) {
  const rows = criticalPartsByDepartment(devices);
  const biggest = Math.max(1, ...rows.flatMap((row) => PARTS.map(({ key }) => row[key] ?? 0)));

  return (
    <Card className="chart-card dv-heat">
      <div className="chart-head">
        <h3>Critical parts by department</h3>
        <p>Bigger dot, more machines. Press a dot to see them.</p>
      </div>
      {rows.length === 0 ? (
        <EmptyState>No devices imported yet.</EmptyState>
      ) : (
        <div className="hm-scroll">
          <table className="hm">
            <thead>
              <tr>
                <th scope="col">Department</th>
                {PARTS.map(({ key, label }) => <th scope="col" key={key}>{label}</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.department}>
                  <th scope="row" title={row.persona}>
                    {row.department} <span className="hm-total">· {row.total}</span>
                  </th>
                  {PARTS.map(({ key, label }) => {
                    const n = row[key] ?? 0;
                    if (n === 0) {
                      return <td key={key}><span className="hm-zero" aria-hidden="true">·</span></td>;
                    }
                    const size = MIN_DOT + Math.round((n / biggest) * (MAX_DOT - MIN_DOT));
                    return (
                      <td key={key}>
                        <button
                          type="button"
                          className="hm-dot"
                          style={{ width: size, height: size }}
                          onClick={() => onSelect?.(row.department, key)}
                          aria-label={`${row.department}: ${n} machines with a critical ${label}. Show them.`}
                        >
                          {n}
                        </button>
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}
