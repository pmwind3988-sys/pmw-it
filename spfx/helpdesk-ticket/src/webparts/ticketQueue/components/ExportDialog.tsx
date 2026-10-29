// The export panel: pick a format and columns, then download exactly the rows the list shows.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { chosenColumns, EXPORT_COLUMNS, ALL_COLUMN_KEYS, toCsv, exportFileName } from '../queue/exportColumns';
import { QueueTicket } from '../queue/ticket';
import { toXlsx } from '../queue/xlsx';
import { useInstanceId, useMountedRef } from '../../helpdeskTicket/components/hooks';
import { countText, ExportContext, exportLines, parseColumns } from './queueFormat';
import { COLUMNS_KEY, readStored, writeStored } from './queueStorage';
import { DownloadIcon } from './queueIcons';
import styles from './TicketQueue.module.scss';

export interface ExportDialogProps {
  /** The filtered, sorted tickets, all pages. */
  rows: QueueTicket[];
  context: ExportContext;
  onClose(): void;
}

type Format = 'xlsx' | 'csv';

const messageOf = (e: unknown): string => (e instanceof Error ? e.message : 'The export did not finish.');

function saveBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function ExportDialog(props: ExportDialogProps): React.ReactElement {
  const { rows, context, onClose } = props;
  const [format, setFormat] = React.useState<Format>('xlsx');
  const [keys, setKeys] = React.useState<string[]>(() => parseColumns(readStored(COLUMNS_KEY)));
  const [working, setWorking] = React.useState(false);
  const [error, setError] = React.useState('');
  const dialog = React.useRef<HTMLDivElement>(null);
  const mounted = useMountedRef();
  const ids = {
    title: useInstanceId('exp-title'),
    text: useInstanceId('exp-text'),
    format: useInstanceId('exp-format')
  };

  // Remember a non-empty column choice; an empty one would come back as "all" anyway.
  React.useEffect(() => {
    if (keys.length) writeStored(COLUMNS_KEY, JSON.stringify(keys));
  }, [keys]);

  // Focus moves into the dialog when it opens.
  React.useEffect(() => {
    dialog.current?.focus();
  }, []);

  const cols = chosenColumns(keys);
  const count = rows.length;

  const toggleKey = (key: string): void => {
    setKeys((k) => (k.indexOf(key) !== -1 ? k.filter((x) => x !== key) : [...k, key]));
  };

  const onKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key !== 'Tab' || !dialog.current) return;
    // Keep Tab inside the dialog while it is open.
    const items = Array.from(
      dialog.current.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled])')
    );
    if (!items.length) return;
    const first = items[0];
    const last = items[items.length - 1];
    const active = document.activeElement;
    if (e.shiftKey && (active === first || active === dialog.current)) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && active === last) {
      e.preventDefault();
      first.focus();
    }
  };

  const download = async (): Promise<void> => {
    if (!cols.length) return;
    setWorking(true);
    setError('');
    try {
      const now = new Date();
      let blob: Blob;
      let ext: 'csv' | 'xlsx';
      if (format === 'csv') {
        blob = new Blob([toCsv(rows, cols)], { type: 'text/csv;charset=utf-8' });
        ext = 'csv';
      } else {
        blob = await toXlsx(rows, cols, { lines: exportLines(context, count, now) });
        ext = 'xlsx';
      }
      saveBlob(blob, exportFileName(context.fileParts, ext, now));
      onClose();
    } catch (e) {
      if (mounted.current) setError(messageOf(e));
    } finally {
      if (mounted.current) setWorking(false);
    }
  };

  return ReactDom.createPortal(
    <div className={styles.scrim}>
      <div
        ref={dialog}
        role="dialog"
        aria-modal="true"
        aria-labelledby={ids.title}
        aria-describedby={ids.text}
        tabIndex={-1}
        className={styles.dialog}
        onKeyDown={onKey}
      >
        <h2 id={ids.title} className={styles.dialogTitle}>
          Export {countText(count)}
        </h2>
        <p id={ids.text} className={styles.muted}>
          Exports the tickets the filters show, in the order shown.
        </p>

        <fieldset className={styles.fieldset}>
          <legend className={styles.fieldLabel}>Format</legend>
          <div className={styles.radios}>
            <label className={styles.radio}>
              <input
                type="radio"
                name={ids.format}
                value="xlsx"
                checked={format === 'xlsx'}
                onChange={() => setFormat('xlsx')}
              />
              Excel (.xlsx)
            </label>
            <label className={styles.radio}>
              <input
                type="radio"
                name={ids.format}
                value="csv"
                checked={format === 'csv'}
                onChange={() => setFormat('csv')}
              />
              CSV
            </label>
          </div>
        </fieldset>

        <fieldset className={styles.fieldset}>
          <legend className={styles.fieldLabel}>Columns</legend>
          <div className={styles.columnTools}>
            <button type="button" className={styles.linkBtn} onClick={() => setKeys(ALL_COLUMN_KEYS.slice())}>
              Select all
            </button>
            <button type="button" className={styles.linkBtn} onClick={() => setKeys([])}>
              Clear
            </button>
            <span className={styles.muted}>
              {cols.length} of {EXPORT_COLUMNS.length} chosen
            </span>
          </div>
          <div className={styles.colGrid}>
            {EXPORT_COLUMNS.map((c) => (
              <label key={c.key} className={styles.checkRow}>
                <input type="checkbox" checked={keys.indexOf(c.key) !== -1} onChange={() => toggleKey(c.key)} />
                <span className={styles.checkName}>{c.header}</span>
              </label>
            ))}
          </div>
        </fieldset>

        {error && (
          <p className={styles.errorBox} role="alert">
            {error}
          </p>
        )}

        <div className={styles.dialogFoot}>
          {!cols.length && (
            <p className={styles.reason} role="status">
              Tick at least one column to download.
            </p>
          )}
          <button type="button" className={styles.secondaryBtn} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className={styles.primaryBtn}
            disabled={!cols.length || working}
            aria-busy={working}
            onClick={() => {
              download().catch(() => undefined);
            }}
          >
            {working ? <span className={styles.spinner} aria-hidden="true" /> : <DownloadIcon />}
            Download
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}
