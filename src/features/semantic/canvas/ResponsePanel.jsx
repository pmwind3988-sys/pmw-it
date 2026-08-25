import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import {
  ChevronLeft, ChevronRight, Download, X, ClipboardList,
} from '../../../components/ui/Icons';
import { maskFor } from '../engine/filterMask.js';
import { countRows, readRows, readRecord, stepRow } from '../engine/rows.js';
import { responseTableColumns } from '../engine/responseFields.js';
import { exportDatasetCsv } from '../export/exporters.js';
import { useSemantic } from '../useSemantic';

const PAGE_SIZE = 50;
const NARROW_QUESTIONS = 5;

/**
 * One response, every field of it.
 *
 * Parked columns are shown and labelled rather than hidden. They are
 * usually the timestamp and who submitted it, which is precisely what
 * somebody who has drilled from a chart into one answer came for.
 */
function ResponseDialog({ dataset, mask, index, onClose, onStep }) {
  const boxRef = useRef(null);
  const fields = useMemo(() => readRecord(dataset, index), [dataset, index]);

  const previous = useMemo(() => stepRow(dataset, mask, index, -1), [dataset, mask, index]);
  const next = useMemo(() => stepRow(dataset, mask, index, 1), [dataset, mask, index]);

  useEffect(() => { boxRef.current?.focus(); }, []);

  const onKeyDown = (event) => {
    if (event.key === 'Escape') {
      // The screen clears the chart selection on Escape from a window
      // listener. Left alone, closing this card would also throw away
      // the selection the user tapped in FROM, and they would be
      // returned to an unfiltered screen.
      event.nativeEvent.stopImmediatePropagation();
      event.preventDefault();
      onClose();
      return;
    }
    if (event.key === 'ArrowLeft' && previous !== null) onStep(previous);
    if (event.key === 'ArrowRight' && next !== null) onStep(next);
  };

  return (
    <div className="sa-record-scrim" role="presentation" onClick={onClose}>
      <div
        className="sa-record"
        role="dialog"
        aria-modal="true"
        aria-label={`Response ${index + 1} in full`}
        tabIndex={-1}
        ref={boxRef}
        onKeyDown={onKeyDown}
        onClick={(e) => e.stopPropagation()}
      >
        <header className="sa-record-head">
          <div>
            <h3 className="sa-record-title">{`Response ${index + 1}`}</h3>
            <p className="sa-record-sub">{`${fields.length} fields, in full`}</p>
          </div>
          <div className="sa-record-nav">
            <button
              type="button"
              aria-label="Previous response"
              disabled={previous === null}
              onClick={() => onStep(previous)}
            >
              <ChevronLeft size={14} />
            </button>
            <button
              type="button"
              aria-label="Next response"
              disabled={next === null}
              onClick={() => onStep(next)}
            >
              <ChevronRight size={14} />
            </button>
            <button type="button" aria-label="Close this response" onClick={onClose}>
              <X size={14} />
            </button>
          </div>
        </header>

        <dl className="sa-record-fields">
          {fields.map((field) => (
            <div
              className={`sa-record-field${field.empty ? ' sa-record-field-empty' : ''}`}
              key={field.name}
            >
              <dt>
                {field.name}
                {field.parked && <span className="sa-badge sa-badge-parked">not charted</span>}
              </dt>
              <dd>{field.empty ? 'Blank' : field.text}</dd>
            </div>
          ))}
        </dl>
      </div>
    </div>
  );
}

/** What the chart tap narrowed the list to, in one sentence. */
function narrowedBy(selection, globalFilters) {
  const parts = [];
  if (selection) parts.push(`${selection.column}: ${selection.values.join(', ')}`);
  for (const filter of globalFilters) {
    parts.push(`${filter.column}: ${(filter.values ?? []).join(', ')}`);
  }
  return parts.join(' · ');
}

/**
 * The responses behind the charts.
 *
 * It reads the SAME mask the charts read — the filter bar plus the
 * chart mark currently tapped — so what is listed here is by
 * construction what the charts above are counting. A separate query
 * would be free to disagree with them, and a list that disagrees with
 * the chart above it is worse than no list.
 *
 * It opens with the sheet, not behind a toggle. Tapping a bar to see
 * who is in it is the main thing this screen is for, and a fold in
 * front of that would have to be opened once per session by every user
 * for no benefit.
 */
export default function ResponsePanel() {
  const { dataset, globalFilters, selection, fileName } = useSemantic();

  const [offset, setOffset] = useState(0);
  const [wide, setWide] = useState(false);
  const [openRow, setOpenRow] = useState(null);

  // No tile id: a selection made on a chart narrows this list. The
  // self-exclusion rule exists so a chart keeps the marks you need to
  // tap again; a list of responses has no such context to protect.
  const mask = useMemo(
    () => (dataset ? maskFor(dataset, globalFilters, selection, null) : null),
    [dataset, globalFilters, selection],
  );

  const total = useMemo(() => countRows(dataset, mask), [dataset, mask]);

  const columns = useMemo(
    () => responseTableColumns(dataset, wide ? null : NARROW_QUESTIONS),
    [dataset, wide],
  );

  // Narrowing the charts renumbers the pages under the user: page three
  // of a set that now holds forty responses would be blank with no
  // explanation. Clamped during render rather than reset from an
  // effect, so there is never a frame showing the empty page.
  const start = offset < total ? offset : 0;

  const rows = useMemo(
    () => (dataset ? readRows(dataset, mask, columns, start, PAGE_SIZE) : []),
    [dataset, mask, columns, start],
  );

  // A response left open while the filters change underneath it would
  // go on showing somebody the charts no longer count.
  const shownRow = openRow !== null && (!mask || mask[openRow]) ? openRow : null;

  const handleExport = useCallback(() => {
    exportDatasetCsv(dataset, `${fileName || 'responses'} responses`, mask);
  }, [dataset, fileName, mask]);

  if (!dataset) return null;

  const narrowed = narrowedBy(selection, globalFilters);
  const hidden = dataset.columns.length - columns.length;
  const shownTo = Math.min(start + rows.length, total);

  return (
    <section className="sa-records" aria-label="The responses behind these charts">
      <header className="sa-records-head">
        <h3 className="sa-records-title">
          <ClipboardList size={15} />
          <span>{narrowed ? 'These responses' : 'Every response'}</span>
          <em>{`${total.toLocaleString()} of ${dataset.rowCount.toLocaleString()}`}</em>
        </h3>

        <div className="sa-records-actions">
          {dataset.columns.length > columns.length || wide ? (
            <Button size="sm" variant="secondary" onClick={() => setWide((v) => !v)}>
              {wide ? 'Fewer questions' : 'All questions'}
            </Button>
          ) : null}
          <Button size="sm" variant="secondary" icon={Download} onClick={handleExport}>
            Export these
          </Button>
        </div>
      </header>

      {/* The list is filtered by something the user tapped somewhere
          above, possibly now scrolled off screen. Saying what, here,
          is what makes the number in the heading legible. */}
      {narrowed && (
        <p className="sa-records-narrowed" role="status">
          {`Narrowed by ${narrowed}`}
        </p>
      )}

      {total === 0 ? (
        <p className="sa-records-empty" role="status">
          Nothing matches. Tap the highlighted mark again, or clear a filter above.
        </p>
      ) : (
        <>
          <div className="sa-table-scroll sa-records-scroll">
            <table className="sa-table sa-records-table">
              <thead>
                <tr>
                  <th className="sa-num">#</th>
                  {columns.map((c) => <th key={c.name}>{c.label ?? c.name}</th>)}
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr
                    key={row.index}
                    className="sa-records-row"
                    tabIndex={0}
                    role="button"
                    aria-label={`Open response ${row.index + 1} in full`}
                    onClick={() => setOpenRow(row.index)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        setOpenRow(row.index);
                      }
                    }}
                  >
                    <td className="sa-num">{row.index + 1}</td>
                    {row.cells.map((cell, i) => (
                      <td key={columns[i].name} title={cell}>
                        {cell === '' ? <span className="sa-records-blank">—</span> : cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="sa-records-foot">
            <span className="sa-summary">
              {`Showing ${(start + 1).toLocaleString()}–${shownTo.toLocaleString()} of ${total.toLocaleString()}`}
              {hidden > 0 && ` · ${hidden} more field${hidden === 1 ? '' : 's'} in each`}
            </span>
            <span className="sa-toolbar-spacer" />
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

      {shownRow !== null && (
        <ResponseDialog
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
