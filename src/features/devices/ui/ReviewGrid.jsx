import { formatMYT } from '../../datastudio/time/malaysiaTime';
import { issuesFor, sortForReview } from '../reviewIssues';

/**
 * `editable: true` marks a column whose value was DERIVED rather than read
 * verbatim out of the file. Those are the only ones worth correcting by hand —
 * editing a parsed value would put the register out of step with the report
 * that produced it.
 */
const COLUMNS = [
  { key: 'computerName', label: 'Computer' },
  { key: 'owner', label: 'Owner', editable: true },
  { key: 'department', label: 'Department', editable: true },
  { key: 'deviceType', label: 'Type', editable: true, options: ['Laptop', 'Desktop', 'Unknown'] },
  { key: 'computerModel', label: 'Model' },
  { key: 'cpuModel', label: 'CPU' },
  { key: 'installedRamGB', label: 'RAM (GB)' },
  { key: 'storageTotalGB', label: 'Storage (GB)' },
  { key: 'storageType', label: 'Disks' },
  { key: 'windowsVersion', label: 'Windows' },
  { key: 'antivirusStatus', label: 'Antivirus' },
  { key: 'riskLevel', label: 'Risk' },
];

export default function ReviewGrid({ devices, excluded, onChange, onToggleRow }) {
  const rows = sortForReview(devices);

  return (
    <div className="rg-wrap">
      <div className="rg-scroll">
        <table className="rg">
          <thead>
            <tr>
              <th className="rg-check"><span className="sr-only">Include</span></th>
              {COLUMNS.map((column) => (
                <th key={column.key}>{column.label}</th>
              ))}
              <th>Scanned</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((device) => {
              const issues = issuesFor(device);
              const id = device.sourceFileName;
              const isExcluded = excluded.has(id);

              const className = [
                issues.length ? 'rg-flagged' : '',
                isExcluded ? 'rg-excluded' : '',
              ].filter(Boolean).join(' ');

              return (
                <tr key={id} className={className || undefined}>
                  <td className="rg-check">
                    <input
                      type="checkbox"
                      checked={!isExcluded}
                      onChange={() => onToggleRow(id)}
                      aria-label={`Include ${device.computerName}`}
                    />
                  </td>

                  {COLUMNS.map((column) => {
                    const value = device[column.key];

                    if (!column.editable) {
                      const riskClass = column.key === 'riskLevel'
                        ? `rg-risk rg-risk-${String(value).toLowerCase()}`
                        : undefined;
                      return <td key={column.key} className={riskClass}>{value ?? '—'}</td>;
                    }

                    return (
                      <td key={column.key} className="rg-editable">
                        {column.options ? (
                          <select
                            value={value ?? 'Unknown'}
                            aria-label={`${column.label} for ${device.computerName}`}
                            onChange={(event) => onChange(id, column.key, event.target.value)}
                          >
                            {column.options.map((option) => (
                              <option key={option} value={option}>{option}</option>
                            ))}
                          </select>
                        ) : (
                          <input
                            type="text"
                            value={value ?? ''}
                            placeholder="—"
                            aria-label={`${column.label} for ${device.computerName}`}
                            onChange={(event) =>
                              onChange(id, column.key, event.target.value || null)}
                          />
                        )}
                      </td>
                    );
                  })}

                  <td title="Malaysia time">{formatMYT(device.scannedOn, 'datetime12')}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <ul className="rg-issues">
        {rows.flatMap((device) =>
          issuesFor(device).map((issue) => (
            <li key={`${device.sourceFileName}-${issue}`}>
              <strong>{device.computerName}</strong> — {issue}
            </li>
          )))}
      </ul>
    </div>
  );
}
