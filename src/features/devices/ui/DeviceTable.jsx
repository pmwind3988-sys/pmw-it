import { useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { Download, Search, X, Pencil, Check, AlertTriangle } from '../../../components/ui/Icons';
import { formatMYT } from '../../datastudio/time/malaysiaTime';
import { formatScalar } from '../formatValue';
import ValueCell from './ValueCell';
import { applyFilters, toCsv } from '../deviceFilters';
import { EDITABLE_FIELDS } from '../sharepoint/updateDevice';
import { DEVICE_COLUMNS } from '../sharepoint/deviceSchema';

/**
 * The columns worth seeing first, in this order. Everything else the scan
 * produced follows in schema order -- the table is deliberately wider than the
 * screen and scrolls sideways, because a field read out of a report and then
 * hidden here is a field nobody knows was collected.
 */
const LEAD_KEYS = [
  'computerName', 'owner', 'department', 'deviceType', 'computerModel', 'cpuModel',
  'installedRamGB', 'storageTotalGB', 'storageType', 'windowsVersion',
  'antivirusStatus', 'riskScore', 'riskLevel',
];

/**
 * `scannedOn` has its own column at the far right, formatted in Malaysia time;
 * `scannedOnMYT` is that same instant again; `rawReport` is the whole report
 * file, one cell of which is taller than the screen.
 */
const SKIP_KEYS = new Set(['scannedOn', 'scannedOnMYT', 'rawReport']);

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

const SCHEMA_COLUMNS = [
  // `Title` holds the computer name, so the schema never lists it.
  { key: 'computerName', label: 'Computer' },
  ...DEVICE_COLUMNS
    .map((column) => ({
      key: keyFor(column.StaticName),
      label: column.Title,
      kind: column.kind,
      numeric: column.kind === 'number',
    }))
    .filter((column) => !SKIP_KEYS.has(column.key)),
];

const rank = (key) => {
  const lead = LEAD_KEYS.indexOf(key);
  return lead === -1 ? LEAD_KEYS.length : lead;
};

// A stable sort, so the columns outside LEAD_KEYS keep their schema order.
const COLUMNS = [...SCHEMA_COLUMNS].sort((a, b) => rank(a.key) - rank(b.key));

/**
 * How many entries of a multi-value field a row shows before collapsing the
 * rest into "+N more". Two is what fits beside the columns either side of it;
 * the device page shows the lot.
 */
const CELL_ENTRIES = 2;

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

const TYPE_OPTIONS = ['Laptop', 'Desktop', 'Unknown'];

export default function DeviceTable({
  devices, filters, onFilterChange, onSave, onDelete, busy,
}) {
  const [sort, setSort] = useState({ key: 'riskScore', dir: 'desc' });
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({});
  const [confirming, setConfirming] = useState(null);

  const startEdit = (device) => {
    setConfirming(null);
    setEditingId(device.id);
    setDraft(Object.fromEntries(EDITABLE_FIELDS.map((f) => [f, device[f] ?? ''])));
  };

  const cancelEdit = () => {
    setEditingId(null);
    setDraft({});
    setConfirming(null);
  };

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
                <th><span className="sr-only">Actions</span></th>
              </tr>
            </thead>
            <tbody>
              {rows.map((device) => {
                const editing = editingId === device.id;
                const manual = new Set(device.manualFields ?? []);

                return (
                  <tr key={device.id ?? device.computerName} className={editing ? 'dt-editing' : undefined}>
                    {COLUMNS.map((column) => {
                      const editable = EDITABLE_FIELDS.includes(column.key);

                      if (editing && editable) {
                        return (
                          <td key={column.key} className="rg-editable">
                            {column.key === 'deviceType' ? (
                              <select
                                value={draft.deviceType ?? 'Unknown'}
                                aria-label={`Type for ${device.computerName}`}
                                onChange={(e) => setDraft((d) => ({ ...d, deviceType: e.target.value }))}
                              >
                                {TYPE_OPTIONS.map((o) => <option key={o} value={o}>{o}</option>)}
                              </select>
                            ) : (
                              <input
                                type="text"
                                value={draft[column.key] ?? ''}
                                placeholder="from the scan"
                                aria-label={`${column.label} for ${device.computerName}`}
                                onChange={(e) => setDraft((d) => ({ ...d, [column.key]: e.target.value }))}
                              />
                            )}
                          </td>
                        );
                      }

                      const risky = column.key === 'riskLevel';
                      return (
                        <td
                          key={column.key}
                          className={risky
                            ? `rg-risk rg-risk-${String(device.riskLevel).toLowerCase()}`
                            : undefined}
                        >
                          {column.key === 'computerName' && device.id != null ? (
                            <Link className="dt-link" to={`/devices/${device.id}`}>
                              {formatScalar(device.computerName, 'text')}
                            </Link>
                          ) : (
                            <ValueCell
                              value={device[column.key]}
                              fieldKey={column.key}
                              kind={column.kind}
                              limit={CELL_ENTRIES}
                            />
                          )}
                          {manual.has(column.key) && (
                            <span className="dt-manual" title="Set by hand — imports leave this alone">
                              edited
                            </span>
                          )}
                        </td>
                      );
                    })}

                    <td title="Malaysia time">{formatMYT(device.scannedOn, 'datetime12')}</td>

                    <td className="dt-actions">
                      {!editing && (
                        <button
                          type="button"
                          className="dt-icon"
                          onClick={() => startEdit(device)}
                          aria-label={`Edit ${device.computerName}`}
                        >
                          <Pencil size={14} />
                        </button>
                      )}

                      {editing && confirming !== device.id && (
                        <>
                          <button
                            type="button"
                            className="dt-icon dt-icon-go"
                            disabled={busy}
                            onClick={async () => { await onSave(device, draft); cancelEdit(); }}
                            aria-label={`Save ${device.computerName}`}
                          >
                            <Check size={14} />
                          </button>
                          <button type="button" className="dt-icon" onClick={cancelEdit}>
                            Cancel
                          </button>
                          <button
                            type="button"
                            className="dt-icon dt-icon-bad"
                            onClick={() => setConfirming(device.id)}
                          >
                            Remove
                          </button>
                        </>
                      )}

                      {confirming === device.id && (
                        <span className="dt-confirm">
                          <AlertTriangle size={14} />
                          Remove {device.computerName}?
                          <button
                            type="button"
                            className="dt-icon dt-icon-bad"
                            disabled={busy}
                            onClick={async () => { await onDelete(device); cancelEdit(); }}
                          >
                            Yes, remove
                          </button>
                          <button type="button" className="dt-icon" onClick={() => setConfirming(null)}>
                            Keep
                          </button>
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}
