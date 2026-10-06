import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { criticalPartsByDepartment } from '../stats/deviceStats';
import { PARTS } from '../standards/defaultStandard';

/**
 * Which department is in the most trouble, and on which part. Each cell is the
 * number of the department's machines Critical on that part; it opens them.
 */
export default function DepartmentHeatmap({ devices, onSelect }) {
  const rows = criticalPartsByDepartment(devices);
  return (
    <Card className="chart-card dv-heat">
      <div className="chart-head">
        <h3>Critical parts by department</h3>
        <p>How many machines in each department are Critical on each part. Click a number to see them.</p>
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
                <th scope="col">Machines</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.department}>
                  <th scope="row" title={row.persona}>{row.department}</th>
                  {PARTS.map(({ key, label }) => (
                    <td key={key}>
                      {row[key] > 0 ? (
                        <button
                          type="button"
                          className="hm-cell g-critical"
                          onClick={() => onSelect?.(row.department, key)}
                          aria-label={`${row.department}: ${row[key]} machines with a critical ${label}. Show them.`}
                        >
                          {row[key]}
                        </button>
                      ) : <span className="hm-zero">0</span>}
                    </td>
                  ))}
                  <td>{row.total}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}
