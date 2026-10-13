import { useState } from 'react';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { countBy } from '../stats/deviceStats';
import { ramBucket, isStale } from '../deviceFilters';

/**
 * The fleet split one way at a time. Every bar is the same neutral brand
 * colour (it is a count, not a grade); machines with no reading go grey.
 * Pressing a bar opens the List filtered to it.
 */
const VIEWS = [
  { id: 'type', label: 'Type', key: 'type', value: (d) => d.deviceType },
  { id: 'windows', label: 'Windows', key: 'windows', value: (d) => d.windowsVersion },
  { id: 'ram', label: 'RAM', key: 'ram', value: (d) => ramBucket(d.installedRamGB) },
  { id: 'storage', label: 'Storage', key: 'storage', value: (d) => d.storageType },
  { id: 'cpu', label: 'CPU age', key: 'cpu', value: (d) => d.cpuAgeBand },
  { id: 'license', label: 'Licence', key: 'license', value: (d) => d.licenseStatus },
  { id: 'persona', label: 'Profile', key: 'persona', value: (d) => d.personaLabel },
];

const NOT_READ = new Set(['Unassigned', 'Unknown']);

export default function FleetMakeup({ devices, onOpen }) {
  const [chosen, setChosen] = useState('windows');
  const view = VIEWS.find((v) => v.id === chosen) ?? VIEWS[1];
  const rows = countBy(devices, view.value);
  const largest = Math.max(1, ...rows.map((row) => row.count));
  const total = devices.length;
  const fresh = total === 0 ? 0 : Math.round(((total - devices.filter((d) => isStale(d)).length) / total) * 100);

  return (
    <Card className="chart-card todo-makeup">
      <div className="chart-head">
        <h3>Fleet makeup</h3>
      </div>
      <div role="group" aria-label="Show the fleet by" className="todo-pills">
        {VIEWS.map((v) => (
          <button
            key={v.id}
            type="button"
            aria-pressed={v.id === view.id}
            className={`todo-pill${v.id === view.id ? ' is-on' : ''}`}
            onClick={() => setChosen(v.id)}
          >
            {v.label}
          </button>
        ))}
      </div>
      {rows.length === 0 ? (
        <EmptyState>No devices imported yet.</EmptyState>
      ) : (
        <ul className="todo-bars">
          {rows.map((row) => {
            const grey = NOT_READ.has(row.label);
            return (
              <li key={row.label}>
                <button
                  type="button"
                  className="todo-bar"
                  onClick={() => onOpen?.(view.key, row.label)}
                >
                  <span className="todo-bar-head">
                    <span>{grey ? 'Not read' : row.label}</span>
                    <strong>{row.count}</strong>
                  </span>
                  <span className="todo-track">
                    <span
                      className={`todo-fill${grey ? ' is-grey' : ''}`}
                      style={{ width: `${Math.round((row.count / largest) * 100)}%` }}
                    />
                  </span>
                </button>
              </li>
            );
          })}
        </ul>
      )}
      {total > 0 && (
        <p className="todo-foot">{fresh}% of machines scanned in the last 6 months.</p>
      )}
    </Card>
  );
}
