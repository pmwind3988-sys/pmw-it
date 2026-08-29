import { useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import Collapsible from '../../../components/ui/Collapsible';
import Pager from '../../../components/ui/Pager';
import { paginate } from '../../../components/ui/paginate';
import {
  Download, Search, X, Pencil, Check, AlertTriangle, Trash2,
} from '../../../components/ui/Icons';
import { formatMYT } from '../../../utils/malaysiaTime';
import { formatScalar } from '../formatValue';
import ValueCell from './ValueCell';
import { applyFilters, toCsv } from '../deviceFilters';
import { EDITABLE_FIELDS } from '../sharepoint/updateDevice';
import { DEVICE_COLUMNS } from '../sharepoint/deviceSchema';
import {
  describeSelection, headerState, isSelectable, selectedDevices,
  toggleAll, toggleId, visibleSelection,
} from '../selection';

/**
 * The columns worth seeing first, in this order. Everything else the scan
 * produced follows in schema order -- the table is deliberately wider than the
 * screen and scrolls sideways, because a field read out of a report and then
 * hidden here is a field nobody knows was collected.
 */
const LEAD_KEYS = [
  'computerName', 'owner', 'department', 'personaLabel', 'fitStatus', 'actionRequired',
  'suggestedFormFactor', 'deviceType', 'licenseStatus', 'computerModel', 'cpuModel',
  'installedRamGB', 'storageTotalGB', 'storageType', 'windowsVersion',
  'antivirusStatus', 'riskScore', 'riskLevel',
];

/**
 * Columns worked out on the way in rather than stored. They are not in the
 * SharePoint schema and never will be — the verdict is recomputed on every read
 * — so the table has to name them itself.
 */
const CALCULATED_COLUMNS = [
  { key: 'personaLabel', label: 'Workload Profile' },
  { key: 'fitStatus', label: 'Device Health' },
  { key: 'actionRequired', label: 'Action Required' },
  { key: 'fitReasons', label: 'Why' },
  { key: 'suggestedFormFactor', label: 'Suggested Form Factor' },
  { key: 'licenseStatus', label: 'Office Licence' },
  { key: 'licenseNote', label: 'Licence Detail' },
  { key: 'gpuClass', label: 'Graphics' },
  { key: 'networkRisk', label: 'Server Link' },
  { key: 'networkNote', label: 'Server Link Detail' },
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
  ...CALCULATED_COLUMNS,
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

const NAME_KEY = 'computerName';
const NAME_COLUMN = COLUMNS.find((column) => column.key === NAME_KEY);

/**
 * The tick box and the computer name share ONE frozen cell rather than two
 * pinned side by side. Two would need the second one's `left` offset to be the
 * exact rendered width of the first, which a table decides for itself and no
 * hard-coded pixel figure can promise. One cell needs no arithmetic, so the
 * pinned edge cannot drift.
 *
 * The scrolling columns therefore skip the name -- the frozen cell has already
 * drawn it. The CSV still exports every column, name included.
 */
const BODY_COLUMNS = COLUMNS.filter((column) => column.key !== NAME_KEY);

/**
 * How many entries of a multi-value field a row shows before collapsing the
 * rest into "+N more". Two is what fits beside the columns either side of it;
 * the device page shows the lot.
 */
const CELL_ENTRIES = 2;

const FILTER_LABELS = {
  risk: 'Risk', attention: 'Needs attention', type: 'Type', department: 'Department',
  os: 'OS', av: 'Antivirus', storage: 'Storage', ram: 'RAM', cpu: 'CPU age',
  windows: 'Windows', stale: 'Stale scans', q: 'Search',
  fit: 'Device health', persona: 'Workload profile', license: 'Office licence',
  server: 'Server link', formfit: 'Form factor mismatch',
};

/**
 * The on/off filters carry no value worth reading -- their chip is the label
 * alone, not "Needs attention: 1".
 */
const FLAG_FILTERS = new Set(['attention', 'stale', 'formfit']);

const chipText = (key, value) =>
  (FLAG_FILTERS.has(key)
    ? (FILTER_LABELS[key] ?? key)
    : `${FILTER_LABELS[key] ?? key}: ${value}`);

function download(name, text) {
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8;' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

const TYPE_OPTIONS = ['Laptop', 'Desktop', 'Unknown'];

/**
 * The two verdict columns are the only coloured ones. They share the risk
 * palette rather than a second one of their own: red is "go and look", amber is
 * "put it on the list", green is "leave it alone", whichever column says it.
 */
const FIT_TONE = {
  Critical: 'critical',
  'Needs Attention': 'watch',
  Optimal: 'ok',
};

const toneClassFor = (device, key) => {
  if (key === 'riskLevel') return `rg-risk rg-risk-${String(device.riskLevel).toLowerCase()}`;
  const tone = key === 'fitStatus' ? FIT_TONE[device.fitStatus] : null;
  return tone ? `rg-risk rg-risk-${tone}` : undefined;
};

export default function DeviceTable({
  devices, filters, onFilterChange, onSave, onDelete, onDeleteMany, busy,
}) {
  const [sort, setSort] = useState({ key: 'riskScore', dir: 'desc' });
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({});
  const [confirming, setConfirming] = useState(null);
  const [selected, setSelected] = useState(() => new Set());
  const [confirmingMany, setConfirmingMany] = useState(false);
  const [removing, setRemoving] = useState(null);
  // Only the page being looked at is laid out. The register runs to a
  // thousand machines and every row of it carries fifty cells.
  const [pageSize, setPageSize] = useState(25);
  const [at, setAt] = useState({ of: '', page: 1 });

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

  /**
   * A tick can never reach off screen. Both readings below are scoped to the
   * rows the filters are showing, and every write is pruned to them as well, so
   * "Remove 3" cannot possibly mean a machine nobody can see — whatever else
   * the stored set happens to be carrying.
   *
   * Scoping on the way in and out beats reconciling the set in an effect: an
   * effect would set state during a render it just caused, and would still have
   * to be right about the same thing.
   *
   * Changing what is ticked also cancels a pending confirm, since it has just
   * changed which machines that confirm was about.
   */
  const updateSelection = (next) => {
    setConfirmingMany(false);
    setSelected((current) => visibleSelection(next(current), rows));
  };

  /**
   * The page number is held together with the list it counts, so a filter that
   * shortens the list puts the reader back on page 1 in the same render. An
   * effect would show page 7 of the old list first and correct it afterwards.
   */
  const of = `${JSON.stringify(filters)}|${sort.key}|${sort.dir}|${pageSize}`;
  const page = at.of === of ? at.page : 1;
  const setPage = (next) => setAt({ of, page: next });
  const paged = useMemo(() => paginate(rows, page, pageSize), [rows, page, pageSize]);

  // Ticks are scoped to the FILTER, not to the page: one ticked on page 1 is
  // still ticked on page 2, and "Remove 12" still cannot mean a machine the
  // filters are hiding. The header box works a page at a time, because that is
  // what somebody pressing it is looking at.
  const chosen = useMemo(() => selectedDevices(selected, rows), [selected, rows]);
  const headBox = headerState(selected, paged.rows);

  const removeChosen = async () => {
    setRemoving({ done: 0, total: chosen.length });
    try {
      await onDeleteMany(chosen, (done, total) => setRemoving({ done, total }));
    } finally {
      // Whatever happened, the ticks go: the rows that went are gone, and the
      // banner names the ones that would not. Leaving them ticked would invite
      // a second attempt at machines that no longer exist.
      setRemoving(null);
      setConfirmingMany(false);
      setSelected(new Set());
    }
  };

  const activeFilters = Object.entries(filters).filter(([, value]) => value);

  const toggleSort = (key) =>
    setSort((current) =>
      (current.key === key
        ? { key, dir: current.dir === 'asc' ? 'desc' : 'asc' }
        : { key, dir: 'desc' }));

  const sortArrow = (key) =>
    (sort.key === key
      ? <span aria-hidden="true">{sort.dir === 'asc' ? ' ▲' : ' ▼'}</span>
      : null);

  let selectionBar;
  if (removing) {
    selectionBar = (
      <span className="dt-bulk-count">
        Removing {removing.done} of {removing.total}…
      </span>
    );
  } else if (confirmingMany) {
    selectionBar = (
      <>
        <AlertTriangle size={14} />
        <span className="dt-bulk-warn">
          Remove {chosen.length === 1 ? 'this device' : `these ${chosen.length} devices`}
          {' '}from the register? {describeSelection(chosen)}
        </span>
        <button
          type="button"
          className="dt-icon dt-icon-bad"
          disabled={busy}
          onClick={removeChosen}
        >
          Yes, remove {chosen.length}
        </button>
        <button type="button" className="dt-icon" onClick={() => setConfirmingMany(false)}>
          Keep them
        </button>
      </>
    );
  } else {
    selectionBar = (
      <>
        <span className="dt-bulk-count">{chosen.length} selected</span>
        <button
          type="button"
          className="dt-icon dt-icon-bad"
          disabled={busy}
          onClick={() => { cancelEdit(); setConfirmingMany(true); }}
        >
          <Trash2 size={14} />
          Remove {chosen.length}
        </button>
        <button type="button" className="dt-icon" onClick={() => updateSelection(() => new Set())}>
          Clear
        </button>
      </>
    );
  }

  return (
    <Card className="rg-card">
      {/* The search box and the filter chips fold away together. On a phone
          they took the top third of the screen before a single machine was
          visible, and most visits here come from a dashboard card that has
          already chosen the filter. What is ON is still said while they are
          folded, because a hidden filter is a list lying about what it shows. */}
      <Collapsible
        id="devices-filters"
        className="dt-fold"
        title="Search and filters"
        defaultOpen={false}
        summary={[
          `${rows.length} of ${devices.length} devices`,
          ...activeFilters.map(([key, value]) => chipText(key, value)),
        ].join(' · ')}
        actions={(
          <Button
            variant="secondary"
            size="sm"
            icon={Download}
            disabled={rows.length === 0}
            onClick={() => download('device-list.csv', toCsv(rows, COLUMNS))}
          >
            CSV
          </Button>
        )}
      >
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
                {chipText(key, value)}
                <X size={12} />
              </button>
            ))}
          </div>
        )}
      </Collapsible>

      {chosen.length > 0 && <div className="dt-bulk">{selectionBar}</div>}

      {rows.length === 0 ? (
        <EmptyState>No devices match these filters.</EmptyState>
      ) : (
        <div className="rg-scroll">
          <table className="rg dt-frozen">
            <thead>
              <tr>
                <th className="dt-identity">
                  <input
                    type="checkbox"
                    className="dt-tick"
                    checked={headBox === 'all'}
                    disabled={busy}
                    ref={(box) => {
                      // Half-ticked is a DOM property, not an attribute, so it
                      // cannot be expressed as a prop.
                      if (box) box.indeterminate = headBox === 'some';
                    }}
                    onChange={() => updateSelection((current) => toggleAll(current, paged.rows))}
                    aria-label="Select every device on this page"
                  />
                  <button
                    type="button"
                    className="dt-sort"
                    onClick={() => toggleSort(NAME_KEY)}
                    aria-label={`Sort by ${NAME_COLUMN.label}`}
                  >
                    {NAME_COLUMN.label}
                    {sortArrow(NAME_KEY)}
                  </button>
                </th>

                {BODY_COLUMNS.map((column) => (
                  <th key={column.key}>
                    <button
                      type="button"
                      className="dt-sort"
                      onClick={() => toggleSort(column.key)}
                      aria-label={`Sort by ${column.label}`}
                    >
                      {column.label}
                      {sortArrow(column.key)}
                    </button>
                  </th>
                ))}
                <th>Scanned</th>
                <th className="dt-actions"><span className="sr-only">Actions</span></th>
              </tr>
            </thead>
            <tbody>
              {paged.rows.map((device) => {
                const editing = editingId === device.id;
                const ticked = isSelectable(device) && selected.has(device.id);
                const manual = new Set(device.manualFields ?? []);

                const rowClass = [
                  editing ? 'dt-editing' : '',
                  ticked ? 'dt-selected' : '',
                ].filter(Boolean).join(' ');

                return (
                  <tr key={device.id ?? device.computerName} className={rowClass || undefined}>
                    <td className="dt-identity">
                      <input
                        type="checkbox"
                        className="dt-tick"
                        checked={ticked}
                        // A row with no id cannot be removed, so offering to
                        // tick it would only lead to an error.
                        disabled={!isSelectable(device) || busy}
                        onChange={() => updateSelection((current) => toggleId(current, device.id))}
                        aria-label={`Select ${device.computerName}`}
                      />
                      {device.id != null ? (
                        <Link className="dt-link" to={`/devices/${device.id}`}>
                          {formatScalar(device.computerName, 'text')}
                        </Link>
                      ) : (
                        formatScalar(device.computerName, 'text')
                      )}
                    </td>

                    {BODY_COLUMNS.map((column) => {
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

                      return (
                        <td key={column.key} className={toneClassFor(device, column.key)}>
                          <ValueCell
                            value={device[column.key]}
                            fieldKey={column.key}
                            kind={column.kind}
                            limit={CELL_ENTRIES}
                          />
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

      {rows.length > 0 && (
        <Pager
          page={paged}
          onPage={setPage}
          size={pageSize}
          onSize={setPageSize}
          label="devices"
        />
      )}
    </Card>
  );
}
