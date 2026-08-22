import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import {
  ChevronLeft, ChevronRight, Download, X, ClipboardList,
} from '../../../components/ui/Icons';
import { maskFor } from '../engine/filterMask.js';
import {
  countRows, readRows, readRecord, stepRow, tableColumns,
} from '../engine/rows.js';
import { exportDatasetCsv } from '../store/exporters.js';
import { useDataStudio } from '../useDataStudio';

const PAGE_SIZE = 50;
const NARROW_COLUMNS = 8;

/**
 * One record, every column of it.
 *
 * Parked columns are shown and labelled rather than hidden. They are
 * usually the timestamp and who submitted it, which is precisely what
 * somebody who has drilled from a chart into one answer came for.
 */
function RecordDialog({ dataset, mask, index, onClose, onStep }) {
  const boxRef = useRef(null);
  const fields = useMemo(() => readRecord(dataset, index), [dataset, index]);

  const previous = useMemo(() => stepRow(dataset, mask, index, -1), [dataset, mask, index]);
  const next = useMemo(() => stepRow(dataset, mask, index, 1), [dataset, mask, index]);

  useEffect(() => { boxRef.current?.focus(); }, []);

  const onKeyDown = (event) => {
    if (event.key === 'Escape') {
      // The canvas clears the cross-filter on Escape from a window
      // listener. Left alone, closing this card would also throw away
      // the selection the user drilled in FROM, and they would be
      // returned to an unfiltered dashboard.
      event.nativeEvent.stopImmediatePropagation();
      event.preventDefault();
      onClose();
      return;
    }
    if (event.key === 'ArrowLeft' && previous !== null) onStep(previous);
    if (event.key === 'ArrowRight' && next !== null) onStep(next);
  };

  return (
    <div className="ds-record-scrim" role="presentation" onClick={onClose}>
      <div
        className="ds-record"
        role="dialog"
        aria-modal="true"
        aria-label={`Row ${index + 1} in full`}
        tabIndex={-1}
        ref={boxRef}
        onKeyDown={onKeyDown}
        onClick={(e) => e.stopPropagation()}
      >
        <header className="ds-record-head">
          <div>
            <h3 className="ds-record-title">{`Row ${index + 1}`}</h3>
            <p className="ds-record-sub">{`${fields.length} columns, as cleaned`}</p>
          </div>
          <div className="ds-record-nav">
            <button
              type="button"
              aria-label="Previous row"
              disabled={previous === null}
              onClick={() => onStep(previous)}
            >
              <ChevronLeft size={14} />
            </button>
            <button
              type="button"
              aria-label="Next row"
              disabled={next === null}
              onClick={() => onStep(next)}
            >
              <ChevronRight size={14} />
            </button>
            <button type="button" aria-label="Close this row" onClick={onClose}>
              <X size={14} />
            </button>
          </div>
        </header>

        <dl className="ds-record-fields">
          {fields.map((field) => (
            <div
              className={`ds-record-field${field.empty ? ' ds-record-field-empty' : ''}`}
              key={field.name}
            >
              <dt>
                {field.name}
                {field.parked && <span className="ds-badge ds-badge-parked">not charted</span>}
              </dt>
              <dd>{field.empty ? 'Blank' : field.text}</dd>
            </div>
          ))}
        </dl>
      </div>
    </div>
  );
}

/**
 * The rows behind the charts (the drill-down half of the canvas).
 *
 * It reads the SAME mask the tiles read -- global filters plus the
 * current cross-filter click -- so what is listed here is by
 * construction what the charts above are counting. A separate query
 * would be free to disagree with them, and a table that disagrees with
 * the chart above it is worse than no table.
 */
export default function RecordsPanel() {
  const { dataset, globalFilters, selection, fileName } = useDataStudio();

  const [open, setOpen] = useState(false);
  const [offset, setOffset] = useState(0);
  const [wide, setWide] = useState(false);
  const [openRow, setOpenRow] = useState(null);

  // No tile id: a selection made on a chart narrows this list. The
  // self-exclusion rule exists so a chart keeps the marks you need to
  // click again; a list of records has no such context to protect.
  const mask = useMemo(
    () => (dataset ? maskFor(dataset, globalFilters, selection, null) : null),
    [dataset, globalFilters, selection],
  );

  const total = useMemo(() => countRows(dataset, mask), [dataset, mask]);

  const columns = useMemo(
    () => tableColumns(dataset, wide ? null : NARROW_COLUMNS),
    [dataset, wide],
  );

  // Narrowing the dashboard renumbers the pages under the user: page
  // three of a set that now holds forty rows would be blank with no
  // explanation. Clamped during render rather than reset from an
  // effect, so there is never a frame showing the empty page.
  const start = offset < total ? offset : 0;

  const rows = useMemo(
    () => (open && dataset ? readRows(dataset, mask, columns, start, PAGE_SIZE) : []),
    [open, dataset, mask, columns, start],
  );

  // A record left open while the filters change underneath it would go
  // on showing a row the dashboard no longer counts.
  const shownRow = openRow !== null && (!mask || mask[openRow]) ? openRow : null;

  const handleExport = useCallback(() => {
    exportDatasetCsv(dataset, `${fileName || 'rows'} rows`, mask);
  }, [dataset, fileName, mask]);

  if (!dataset) return null;

  const hidden = dataset.columns.length - columns.length;
  const shownTo = Math.min(start + rows.length, total);

  return (
    <section className="ds-records" aria-label="The rows behind these charts">
      <header className="ds-records-head">
        <button
          type="button"
          className="ds-records-toggle"
          aria-expanded={open}
          onClick={() => setOpen((v) => !v)}
        >
          <ClipboardList size={15} />
          <span>{open ? 'Hide the rows' : 'See the rows behind these charts'}</span>
          <em>{`${total.toLocaleString()} row${total === 1 ? '' : 's'}`}</em>
        </button>

        {open && (
          <div className="ds-records-actions">
            {dataset.columns.length > NARROW_COLUMNS && (
              <Button size="sm" variant="secondary" onClick={() => setWide((v) => !v)}>
                {wide ? 'Show fewer columns' : `Show all ${dataset.columns.length} columns`}
              </Button>
            )}
            <Button size="sm" variant="secondary" icon={Download} onClick={handleExport}>
              Export these rows
            </Button>
          </div>
        )}
      </header>

      {open && (
        <>
          {total === 0 ? (
            <p className="ds-records-empty" role="status">
              No rows match the current filters. Clear a filter above to see records again.
            </p>
          ) : (
            <>
              <div className="ds-table-scroll ds-records-scroll">
                <table className="ds-table ds-records-table">
                  <thead>
                    <tr>
                      <th className="ds-num">#</th>
                      {columns.map((c) => <th key={c.name}>{c.name}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row) => (
                      <tr
                        key={row.index}
                        className="ds-records-row"
                        tabIndex={0}
                        role="button"
                        aria-label={`Open row ${row.index + 1} in full`}
                        onClick={() => setOpenRow(row.index)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault();
                            setOpenRow(row.index);
                          }
                        }}
                      >
                        <td className="ds-num">{row.index + 1}</td>
                        {row.cells.map((cell, i) => (
                          <td key={columns[i].name} title={cell}>
                            {cell === '' ? <span className="ds-records-blank">—</span> : cell}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="ds-records-foot">
                <span className="ds-summary">
                  {`Showing ${(start + 1).toLocaleString()}–${shownTo.toLocaleString()} of ${total.toLocaleString()}`}
                  {hidden > 0 && ` · ${hidden} more column${hidden === 1 ? '' : 's'} in each row`}
                </span>
                <span className="ds-toolbar-spacer" />
                <Button
                  size="sm"
                  variant="secondary"
                  icon={ChevronLeft}
                  disabled={start === 0}
                  onClick={() => setOffset(Math.max(0, start - PAGE_SIZE))}
                >
                  Previous
                </Button>
                <Button
                  size="sm"
                  variant="secondary"
                  icon={ChevronRight}
                  disabled={start + PAGE_SIZE >= total}
                  onClick={() => setOffset(start + PAGE_SIZE)}
                >
                  Next
                </Button>
              </div>
            </>
          )}
        </>
      )}

      {shownRow !== null && (
        <RecordDialog
          dataset={dataset}
          mask={mask}
          index={shownRow}
          onClose={() => setOpenRow(null)}
          onStep={setOpenRow}
        />
      )}
    </section>
  );
}
