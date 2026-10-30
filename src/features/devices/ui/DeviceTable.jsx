import { Fragment, useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import Spinner from '../../../components/ui/Spinner';
import Pager from '../../../components/ui/Pager';
import { paginate } from '../../../components/ui/paginate';
import {
  Download, Search, X, Pencil, Check, AlertTriangle, Trash2, MoreHorizontal,
} from '../../../components/ui/Icons';
import { formatMYT } from '../../../utils/malaysiaTime';
import { formatScalar } from '../formatValue';
import ValueCell from './ValueCell';
import FilterMenu from './FilterMenu';
import ColumnMenu from './ColumnMenu';
import { applyFilters, IN_FLEET, toCsv } from '../deviceFilters';
import { buildFilterGroups, chipLabel } from '../filterOptions';
import {
  ALL_COLUMNS, DEFAULT_SET, columnsForKeys, keysForSet, parseColumnSet, toggleColumn, valueOfColumn,
} from '../columnSets';
import { gradeRank, securityRiskOf, worstGradeOf } from '../health';
import { EDITABLE_FIELDS } from '../sharepoint/updateDevice';
import { GRADE_SLUG } from '../standards/gradeColors';
import { GRADES, UNKNOWN } from '../standards/defaultStandard';
import {
  describeSelection, headerState, isSelectable, selectedDevices,
  toggleAll, toggleId, visibleSelection,
} from '../selection';

const NAME_KEY = 'computerName';
const NAME_COLUMN = ALL_COLUMNS.find((column) => column.key === NAME_KEY);

/**
 * How many entries of a multi-value field a row shows before collapsing the
 * rest into "+N more". Two is what fits beside the columns either side of it;
 * the device page shows the lot.
 */
const CELL_ENTRIES = 2;

const TYPE_OPTIONS = ['Laptop', 'Desktop', 'Unknown'];
const NO_LOCATIONS = [];

const GRADE_KEYS = new Set(['gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows']);

// Severity order, worst first, so sorting a grade column groups the machines that need work.
const GRADE_ORDER = [...GRADES, UNKNOWN];
const severityOf = (grade) => {
  const at = GRADE_ORDER.indexOf(grade);
  return at === -1 ? GRADE_ORDER.length : at;
};

const DEFAULT_SORT = { key: 'worst', dir: 'asc' };

/** The sort dropdown's choices. A column-header sort that matches none of them shows as "Column order". */
const SORT_PRESETS = [
  { id: 'worst', label: 'Needs work first', sort: DEFAULT_SORT },
  { id: 'name', label: 'Computer A–Z', sort: { key: NAME_KEY, dir: 'asc' } },
  { id: 'scanned', label: 'Newest scan', sort: { key: 'scannedOn', dir: 'desc' } },
];

const presetFor = (sort) =>
  SORT_PRESETS.find((p) => p.sort.key === sort.key && p.sort.dir === sort.dir)?.id ?? 'custom';

const sortLabelOf = (sort) =>
  (sort.key === 'worst'
    ? 'Needs work first'
    : ALL_COLUMNS.find((column) => column.key === sort.key)?.label ?? sort.key);

const COLUMN_STORAGE_KEY = 'deviceColumns';

/** The stored choice, or Essentials. Storage can be blocked, and then the choice is simply not remembered. */
function readColumnSet() {
  try {
    return parseColumnSet(window.localStorage.getItem(COLUMN_STORAGE_KEY));
  } catch {
    return DEFAULT_SET;
  }
}

function writeColumnSet(id) {
  try {
    window.localStorage.setItem(COLUMN_STORAGE_KEY, id);
  } catch {
    // A private window cannot remember it; the table still works.
  }
}

/** Epoch milliseconds of a scan, or null when it has none. */
function scannedAt(device) {
  if (device.scannedOn === null || device.scannedOn === undefined || device.scannedOn === '') return null;
  const time = new Date(device.scannedOn).getTime();
  return Number.isFinite(time) ? time : null;
}

function scannedText(device) {
  try {
    return formatMYT(device.scannedOn, 'datetime12');
  } catch {
    return '—';
  }
}

function valueForSort(device, key) {
  if (key === 'worst') return gradeRank(worstGradeOf(device));
  if (key === 'scannedOn') return scannedAt(device);
  return device[key];
}

const byName = (a, b) =>
  String(a.computerName ?? '').localeCompare(String(b.computerName ?? ''));

function download(name, text) {
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8;' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

const GRADE_CLASS = (grade) => `pc-chip g-${GRADE_SLUG[grade] ?? 'unknown'}`;

export default function DeviceTable({
  devices, filters, onFilterChange, onSave, onDelete, onDeleteMany, onRename, busy,
  locationOptions = NO_LOCATIONS,
}) {
  // Removing is admin-only: without the handlers there are no ticks, no
  // Remove in the row menu and no bulk bar, rather than controls that fail.
  const canRemove = Boolean(onDelete && onDeleteMany);
  const [sort, setSort] = useState(DEFAULT_SORT);
  const [columnKeys, setColumnKeys] = useState(() => keysForSet(readColumnSet()));
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({});
  // The row key (id, or name for a row with none) whose remove is awaiting a yes.
  const [confirming, setConfirming] = useState(null);
  // The row whose ⋯ menu is open.
  const [menuFor, setMenuFor] = useState(null);
  const [menuOpen, setMenuOpen] = useState(false);
  const [menuFocus, setMenuFocus] = useState(null);
  const [selected, setSelected] = useState(() => new Set());
  const [confirmingMany, setConfirmingMany] = useState(false);
  const [removing, setRemoving] = useState(null);
  // Only the page being looked at is laid out. The register runs to a
  // thousand machines and every row of it carries fifty cells.
  const [pageSize, setPageSize] = useState(25);
  const [at, setAt] = useState({ of: '', page: 1 });

  const columns = useMemo(() => columnsForKeys(columnKeys), [columnKeys]);

  // The status filter has ONE default: In use or in repair. Everything else
  // the menu can ask for is an explicit choice, and 'All' lifts the status
  // filter altogether.
  const effectiveFilters = useMemo(() => ({
    ...filters,
    status: filters.status === 'All' ? '' : (filters.status || IN_FLEET),
  }), [filters]);

  const groups = useMemo(
    () => buildFilterGroups(devices, locationOptions),
    [devices, locationOptions],
  );

  // How many machines the typed program name keeps, under the same status
  // filter the rows use, so the figure in the menu is the List's own count.
  const softwareCount = useMemo(() => (filters.software
    ? applyFilters(devices, { status: effectiveFilters.status, software: filters.software }).length
    : null), [devices, filters.software, effectiveFilters.status]);

  // A row's menu and its confirm are keyed by id, or by name for a row with none.
  const rowKeyOf = (device) => device.id ?? device.computerName;

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
    const filtered = applyFilters(devices, effectiveFilters);
    const column = ALL_COLUMNS.find((c) => c.key === sort.key);
    const numeric = sort.key === 'worst' || sort.key === 'scannedOn' || Boolean(column?.numeric);

    return [...filtered].sort((a, b) => {
      const left = valueForSort(a, sort.key);
      const right = valueForSort(b, sort.key);

      // Nulls sink to the bottom whichever way the column is sorted — a blank
      // is never the "most" or "least" of anything.
      if (left === null || left === undefined) return right === null || right === undefined ? byName(a, b) : 1;
      if (right === null || right === undefined) return -1;

      let compared;
      if (GRADE_KEYS.has(sort.key)) compared = severityOf(left) - severityOf(right);
      else if (numeric) compared = left - right;
      else compared = String(left).localeCompare(String(right));

      // Machines that tie keep a stable, readable order.
      if (compared === 0) return byName(a, b);
      return sort.dir === 'asc' ? compared : -compared;
    });
  }, [devices, effectiveFilters, sort]);

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
  const of = `${JSON.stringify(effectiveFilters)}|${sort.key}|${sort.dir}|${pageSize}`;
  const page = at.of === of ? at.page : 1;
  const setPage = (next) => setAt({ of, page: next });
  const paged = useMemo(() => paginate(rows, page, pageSize), [rows, page, pageSize]);

  // Ticks are scoped to the FILTER, not to the page: one ticked on page 1 is
  // still ticked on page 2, and "Remove 12" still cannot mean a machine the
  // filters are hiding. The header box works a page at a time, because that is
  // what somebody pressing it is looking at.
  const chosen = useMemo(() => selectedDevices(selected, rows), [selected, rows]);
  const headBox = headerState(selected, paged.rows);

  // Closes a row's ⋯ menu on an outside press or Escape. Listeners exist only while one is open.
  useEffect(() => {
    if (menuFor === null) return undefined;

    const onPointer = (event) => {
      if (!event.target.closest?.('.dt-rowmenu')) setMenuFor(null);
    };
    const onKey = (event) => {
      if (event.key === 'Escape') setMenuFor(null);
    };
    document.addEventListener('pointerdown', onPointer);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onPointer);
      document.removeEventListener('keydown', onKey);
    };
  }, [menuFor]);

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

  const activeFilters = Object.entries(filters).filter(([key, value]) =>
    value && key !== 'q' && key !== 'status');
  const hasStatusFilter = Boolean(filters.status);
  const anyFilter = activeFilters.length > 0 || hasStatusFilter || Boolean(filters.q);

  // One call for every key: separate calls in a row would each start from the
  // same address and undo one another.
  const clearAll = () => {
    onFilterChange(Object.fromEntries(
      Object.entries(filters).filter(([, value]) => value).map(([key]) => [key, '']),
    ));
  };

  const openMenu = (focusFilter = null) => {
    setMenuFocus(focusFilter);
    setMenuOpen(true);
  };

  const pick = (filter, value) => {
    setMenuOpen(false);
    onFilterChange(filter, value);
  };

  const toggleSort = (key) =>
    setSort((current) =>
      (current.key === key
        ? { key, dir: current.dir === 'asc' ? 'desc' : 'asc' }
        : { key, dir: 'desc' }));

  const sortArrow = (key) =>
    (sort.key === key
      ? <span aria-hidden="true">{sort.dir === 'asc' ? ' ▲' : ' ▼'}</span>
      : null);

  // A set is remembered; ticking columns on top of it is not, so a refresh
  // puts the list back on the last set chosen.
  const chooseColumnSet = (next) => {
    writeColumnSet(next);
    setColumnKeys(keysForSet(next));
  };

  const chooseSortPreset = (id) => {
    const preset = SORT_PRESETS.find((p) => p.id === id);
    if (preset) setSort(preset.sort);
  };

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
          aria-busy={busy || undefined}
          onClick={removeChosen}
        >
          {busy && <Spinner size={12} />}
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

  const statusValue = filters.status || IN_FLEET;

  return (
    <Card className="rg-card dl-card">
      <div className="dl-search">
        <Search size={16} />
        <input
          type="search"
          value={filters.q ?? ''}
          placeholder="Search name, owner, serial or AnyDesk ID"
          aria-label="Search name, owner, serial or AnyDesk ID"
          onChange={(event) => onFilterChange('q', event.target.value)}
        />
      </div>

      <div className="dl-filterbar">
        <strong className="dl-count">
          {rows.length} <span>of {devices.length}</span>
        </strong>

        <span className="dt-chip dt-chip-status">
          <button
            type="button"
            className="dt-chip-main"
            aria-expanded={menuOpen}
            onClick={() => openMenu('status')}
          >
            {chipLabel('status', statusValue)}
            <span aria-hidden="true"> ▾</span>
          </button>
          {hasStatusFilter && (
            <button
              type="button"
              className="dt-chip-x"
              aria-label="Go back to the default status filter"
              onClick={() => onFilterChange('status', '')}
            >
              <X size={12} />
            </button>
          )}
        </span>

        {activeFilters.map(([key, value]) => (
          <span className="dt-chip" key={key}>
            <span className="dt-chip-main">{chipLabel(key, value)}</span>
            <button
              type="button"
              className="dt-chip-x"
              onClick={() => onFilterChange(key, '')}
              aria-label={`Remove the ${chipLabel(key, value)} filter`}
            >
              <X size={12} />
            </button>
          </span>
        ))}

        <FilterMenu
          open={menuOpen}
          onOpenChange={setMenuOpen}
          focusFilter={menuFocus}
          software={{
            value: filters.software ?? '',
            count: softwareCount,
            onChange: (value) => onFilterChange('software', value),
          }}
          groups={groups}
          onPick={pick}
        />

        {anyFilter && (
          <button type="button" className="dl-clear" onClick={clearAll}>
            Clear all
          </button>
        )}

        <span className="dl-spacer" />

        <label className="dl-select">
          Sort
          <select
            value={presetFor(sort)}
            onChange={(event) => chooseSortPreset(event.target.value)}
          >
            {SORT_PRESETS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
            {presetFor(sort) === 'custom' && <option value="custom">Column order</option>}
          </select>
        </label>

        <ColumnMenu
          keys={columnKeys}
          onChooseSet={chooseColumnSet}
          onToggle={(key) => setColumnKeys((current) => toggleColumn(current, key))}
        />

        <Button
          variant="secondary"
          size="sm"
          icon={Download}
          disabled={rows.length === 0}
          onClick={() => download(
            'device-list.csv',
            toCsv(rows.map((device) => ({ ...device, scannedOn: scannedText(device) })), ALL_COLUMNS),
          )}
        >
          CSV
        </Button>
      </div>

      {canRemove && chosen.length > 0 && <div className="dt-bulk">{selectionBar}</div>}

      {rows.length === 0 ? (
        <EmptyState>No devices match these filters.</EmptyState>
      ) : (
        <div className="rg-scroll">
          <table className="rg dt-frozen">
            <thead>
              <tr>
                <th className="dt-identity">
                  {canRemove && (
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
                  )}
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

                {columns.map((column) => (
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
              </tr>
            </thead>
            <tbody>
              {paged.rows.map((device, rowIndex) => {
                const rowKey = rowKeyOf(device);
                const editing = editingId === device.id && device.id != null;
                const ticked = isSelectable(device) && selected.has(device.id);
                const menuOpenHere = menuFor === rowKey;
                const manual = new Set(device.manualFields ?? []);
                const name = formatScalar(device.computerName, 'text');

                const rowClass = [
                  editing ? 'dt-editing' : '',
                  ticked ? 'dt-selected' : '',
                ].filter(Boolean).join(' ');

                return (
                  <Fragment key={rowKey}>
                    <tr className={rowClass || undefined}>
                      <td className="dt-identity">
                        <div className="dt-name">
                          {canRemove && (
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
                          )}
                          {device.id != null ? (
                            <Link className="dt-link" to={`/devices/${device.id}`}>{name}</Link>
                          ) : (
                            <span className="dt-link">{name}</span>
                          )}
                          {device.previousNames?.length > 0 && (
                            <span className="dt-was" title={`Previous names: ${device.previousNames.join(', ')}`}>
                              was {device.previousNames.at(-1)}
                            </span>
                          )}

                          <div className="dt-rowmenu">
                            <button
                              type="button"
                              className="dt-more"
                              aria-label={`Actions for ${device.computerName}`}
                              aria-haspopup="menu"
                              aria-expanded={menuOpenHere}
                              onClick={() => setMenuFor(menuOpenHere ? null : rowKey)}
                            >
                              <MoreHorizontal size={16} />
                            </button>

                            {/* The last rows open upward, or the table's own scroll box cuts the menu off. */}
                            {menuOpenHere && (
                              <div className={`dt-menu${rowIndex >= paged.rows.length - 3 && rowIndex > 2 ? ' dt-menu-up' : ''}`} role="menu">
                                {device.id != null && (
                                  <Link
                                    role="menuitem"
                                    className="dt-menu-item"
                                    to={`/devices/${device.id}`}
                                    onClick={() => setMenuFor(null)}
                                  >
                                    Open
                                  </Link>
                                )}
                                <button
                                  type="button"
                                  role="menuitem"
                                  className="dt-menu-item"
                                  onClick={() => { setMenuFor(null); startEdit(device); }}
                                >
                                  <Pencil size={14} />
                                  Edit details
                                </button>
                                {/* Admin only, like Remove; a row with no id cannot be written to. */}
                                {onRename && device.id != null && (
                                  <button
                                    type="button"
                                    role="menuitem"
                                    className="dt-menu-item"
                                    onClick={() => { setMenuFor(null); onRename(device); }}
                                  >
                                    Rename…
                                  </button>
                                )}
                                {canRemove && (
                                  <button
                                    type="button"
                                    role="menuitem"
                                    className="dt-menu-item dt-menu-bad"
                                    onClick={() => { setMenuFor(null); setConfirming(rowKey); }}
                                  >
                                    Remove…
                                  </button>
                                )}
                              </div>
                            )}
                          </div>
                        </div>

                        {editing && (
                          <div className="dt-edit-actions">
                            <button
                              type="button"
                              className="dt-icon dt-icon-go"
                              disabled={busy}
                              aria-busy={busy || undefined}
                              onClick={async () => { await onSave(device, draft); cancelEdit(); }}
                              aria-label={`Save ${device.computerName}`}
                            >
                              {busy ? <Spinner size={12} /> : <Check size={14} />}
                              Save
                            </button>
                            <button type="button" className="dt-icon" onClick={cancelEdit}>
                              Cancel
                            </button>
                          </div>
                        )}
                      </td>

                      {columns.map((column) => {
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
                                  placeholder="blank = use the scan"
                                  aria-label={`${column.label} for ${device.computerName}`}
                                  onChange={(e) => setDraft((d) => ({ ...d, [column.key]: e.target.value }))}
                                />
                              )}
                            </td>
                          );
                        }

                        let content;
                        if (GRADE_KEYS.has(column.key)) {
                          content = (
                            <span className={GRADE_CLASS(device[column.key])}>
                              {device[column.key] ?? UNKNOWN}
                            </span>
                          );
                        } else if (column.key === 'scannedOn') {
                          content = scannedText(device);
                        } else if (column.key === 'riskLevel' && device.riskLevel) {
                          content = (
                            <span className={`rg-risk rg-risk-${String(device.riskLevel).toLowerCase()}`}>
                              {securityRiskOf(device.riskLevel)}
                            </span>
                          );
                        } else {
                          content = (
                            <ValueCell
                              value={valueOfColumn(column, device)}
                              fieldKey={column.key}
                              kind={column.kind}
                              limit={CELL_ENTRIES}
                            />
                          );
                        }

                        return (
                          <td key={column.key} className={column.key === 'scannedOn' ? 'dt-when' : undefined}>
                            {content}
                            {manual.has(column.key) && (
                              <span className="dt-manual" title="Set by hand — imports leave this alone">
                                edited
                              </span>
                            )}
                          </td>
                        );
                      })}
                    </tr>

                    {canRemove && confirming === rowKey && (
                      <tr className="dt-confirm-row">
                        <td colSpan={columns.length + 1}>
                          <div className="dt-confirm">
                            <AlertTriangle size={14} />
                            Remove {device.computerName}?
                            <button
                              type="button"
                              className="dt-icon dt-icon-bad"
                              disabled={busy}
                              aria-busy={busy || undefined}
                              onClick={async () => { await onDelete(device); cancelEdit(); }}
                            >
                              {busy && <Spinner size={12} />}
                              Yes, remove
                            </button>
                            <button type="button" className="dt-icon" onClick={() => setConfirming(null)}>
                              Keep
                            </button>
                          </div>
                        </td>
                      </tr>
                    )}
                  </Fragment>
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
          label={`devices · sorted: ${sortLabelOf(sort)}`}
        />
      )}
    </Card>
  );
}
