import { useMemo, useState } from 'react';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { Download, Search, X } from '../../../components/ui/Icons';
import { formatMYT } from '../../datastudio/time/malaysiaTime';
import { applyFilters, toCsv } from '../deviceFilters';

const COLUMNS = [
  { key: 'computerName', label: 'Computer' },
  { key: 'owner', label: 'Owner' },
  { key: 'department', label: 'Department' },
  { key: 'deviceType', label: 'Type' },
  { key: 'computerModel', label: 'Model' },
  { key: 'cpuModel', label: 'CPU' },
  { key: 'installedRamGB', label: 'RAM (GB)', numeric: true },
  { key: 'storageTotalGB', label: 'Storage (GB)', numeric: true },
  { key: 'storageType', label: 'Disks' },
  { key: 'windowsVersion', label: 'Windows' },
  { key: 'antivirusStatus', label: 'Antivirus' },
  { key: 'riskScore', label: 'Score', numeric: true },
  { key: 'riskLevel', label: 'Risk' },
];

const FILTER_LABELS = {
  risk: 'Risk', type: 'Type', department: 'Department', os: 'OS', av: 'Antivirus',
  storage: 'Storage', ram: 'RAM', cpu: 'CPU age', windows: 'Windows', stale: 'Stale scans',
  q: 'Search',
};

function download(name, text) {
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8;' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

export default function DeviceTable({ devices, filters, onFilterChange }) {
  const [sort, setSort] = useState({ key: 'riskScore', dir: 'desc' });

  const rows = useMemo(() => {
    const filtered = applyFilters(devices, filters);
    const column = COLUMNS.find((c) => c.key === sort.key);

    return [...filtered].sort((a, b) => {
      const left = a[sort.key];
      const right = b[sort.key];

      // Nulls sink to the bottom whichever way the column is sorted — a blank
      // is never the "most" or "least" of anything.
      if (left === null || left === undefined) return 1;
      if (right === null || right === undefined) return -1;

      const compared = column?.numeric
        ? left - right
        : String(left).localeCompare(String(right));
      return sort.dir === 'asc' ? compared : -compared;
    });
  }, [devices, filters, sort]);

  const activeFilters = Object.entries(filters).filter(([, value]) => value);

  const toggleSort = (key) =>
    setSort((current) =>
      (current.key === key
        ? { key, dir: current.dir === 'asc' ? 'desc' : 'asc' }
        : { key, dir: 'desc' }));

  return (
    <Card>
      <div className="dt-head">
        <div className="dt-search">
          <Search size={14} />
          <input
            type="search"
            value={filters.q ?? ''}
            placeholder="Search computer or owner"
            aria-label="Search computer or owner"
            onChange={(event) => onFilterChange('q', event.target.value)}
          />
        </div>

        <div className="dt-head-right">
          <span className="dt-count">{rows.length} of {devices.length}</span>
          <Button
            variant="secondary"
            size="sm"
            icon={Download}
            disabled={rows.length === 0}
            onClick={() => download('device-list.csv', toCsv(rows, COLUMNS))}
          >
            CSV
          </Button>
        </div>
      </div>

      {activeFilters.length > 0 && (
        <div className="dt-chips">
          {activeFilters.map(([key, value]) => (
            <button
              type="button"
              className="dt-chip"
              key={key}
              onClick={() => onFilterChange(key, '')}
              aria-label={`Remove the ${FILTER_LABELS[key] ?? key} filter`}
            >
              {FILTER_LABELS[key] ?? key}: {value}
              <X size={12} />
            </button>
          ))}
        </div>
      )}

      {rows.length === 0 ? (
        <EmptyState>No devices match these filters.</EmptyState>
      ) : (
        <div className="rg-scroll">
          <table className="rg">
            <thead>
              <tr>
                {COLUMNS.map((column) => (
                  <th key={column.key}>
                    <button
                      type="button"
                      className="dt-sort"
                      onClick={() => toggleSort(column.key)}
                      aria-label={`Sort by ${column.label}`}
                    >
                      {column.label}
                      {sort.key === column.key && (
                        <span aria-hidden="true">{sort.dir === 'asc' ? ' ▲' : ' ▼'}</span>
                      )}
                    </button>
                  </th>
                ))}
                <th>Scanned</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((device) => (
                <tr key={device.id ?? device.computerName}>
                  {COLUMNS.map((column) => (
                    <td
                      key={column.key}
                      className={column.key === 'riskLevel'
                        ? `rg-risk rg-risk-${String(device.riskLevel).toLowerCase()}`
                        : undefined}
                    >
                      {device[column.key] ?? '—'}
                    </td>
                  ))}
                  <td title="Malaysia time">{formatMYT(device.scannedOn, 'datetime12')}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}
