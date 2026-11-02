// Pick the table's columns: drag them into order (or use the arrows, which
// also work on a phone and from the keyboard), remove one, add one back, or
// return to the default. Every change applies at once and is remembered in
// this browser.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  addColumn, columnOf, hiddenColumns, isDefault, LOCKED_COLUMN, DEFAULT_COLUMNS, moveColumn, removeColumn
} from '../queue/tableColumns';
import { useInstanceId } from '../../helpdeskTicket/components/hooks';
import styles from './TicketQueue.module.scss';

export interface ColumnsDialogProps {
  keys: string[];
  onChange(keys: string[]): void;
  onClose(): void;
}

const svg = {
  width: 16,
  height: 16,
  viewBox: '0 0 24 24',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 2,
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true,
  focusable: false
};

const Grip = (): React.ReactElement => (
  <svg {...svg}>
    <circle cx="9" cy="6" r="1" />
    <circle cx="15" cy="6" r="1" />
    <circle cx="9" cy="12" r="1" />
    <circle cx="15" cy="12" r="1" />
    <circle cx="9" cy="18" r="1" />
    <circle cx="15" cy="18" r="1" />
  </svg>
);

const Up = (): React.ReactElement => (
  <svg {...svg}>
    <path d="m6 15 6-6 6 6" />
  </svg>
);

const Down = (): React.ReactElement => (
  <svg {...svg}>
    <path d="m6 9 6 6 6-6" />
  </svg>
);

const Cross = (): React.ReactElement => (
  <svg {...svg}>
    <path d="M6 6l12 12" />
    <path d="M18 6 6 18" />
  </svg>
);

const Plus = (): React.ReactElement => (
  <svg {...svg}>
    <path d="M12 5v14" />
    <path d="M5 12h14" />
  </svg>
);

export function ColumnsDialog(props: ColumnsDialogProps): React.ReactElement {
  const { keys, onChange, onClose } = props;
  const [from, setFrom] = React.useState<number | undefined>(undefined);
  const [over, setOver] = React.useState<number | undefined>(undefined);
  const dialog = React.useRef<HTMLDivElement>(null);
  const ids = { title: useInstanceId('cols-title'), text: useInstanceId('cols-text') };

  React.useEffect(() => {
    dialog.current?.focus();
  }, []);

  const onKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key !== 'Tab' || !dialog.current) return;
    // Keep Tab inside the dialog while it is open.
    const items = Array.from(dialog.current.querySelectorAll<HTMLElement>('button:not([disabled])'));
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

  const endDrag = (): void => {
    setFrom(undefined);
    setOver(undefined);
  };

  const hidden = hiddenColumns(keys);

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
          Columns
        </h2>
        <p id={ids.text} className={styles.muted}>
          Drag a column, or use the arrows, to change its place. The top of the list is the left of the table.
          Changes show straight away and are remembered in this browser.
        </p>

        <div className={styles.fieldset}>
          <h3 className={styles.fieldLabel}>{`Shown (${keys.length})`}</h3>
          <ol className={styles.colList}>
            {keys.map((key, i) => {
              const col = columnOf(key);
              if (!col) return null;
              const locked = key === LOCKED_COLUMN;
              const cls = [
                styles.colItem,
                from === i ? styles.colDragging : '',
                over === i && from !== undefined && from !== i ? (from < i ? styles.colOverAfter : styles.colOverBefore) : ''
              ].join(' ');
              return (
                <li
                  key={key}
                  className={cls}
                  draggable={!locked}
                  onDragStart={(e) => {
                    setFrom(i);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', key);
                  }}
                  onDragOver={(e) => {
                    if (from === undefined || i === 0) return;
                    e.preventDefault();
                    if (over !== i) setOver(i);
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (from !== undefined && over !== undefined) onChange(moveColumn(keys, from, over));
                    endDrag();
                  }}
                  onDragEnd={endDrag}
                >
                  <span className={locked ? styles.gripOff : styles.grip}>
                    <Grip />
                  </span>
                  <span className={styles.colName}>
                    {col.label}
                    {locked && <span className={styles.colNote}>Always first: it links to the ticket</span>}
                  </span>
                  {!locked && (
                    <span className={styles.colActions}>
                      <button
                        type="button"
                        className={styles.smallBtn}
                        aria-label={`Move ${col.label} up`}
                        title="Move up"
                        disabled={i <= 1}
                        onClick={() => onChange(moveColumn(keys, i, i - 1))}
                      >
                        <Up />
                      </button>
                      <button
                        type="button"
                        className={styles.smallBtn}
                        aria-label={`Move ${col.label} down`}
                        title="Move down"
                        disabled={i >= keys.length - 1}
                        onClick={() => onChange(moveColumn(keys, i, i + 1))}
                      >
                        <Down />
                      </button>
                      <button
                        type="button"
                        className={styles.smallBtn}
                        aria-label={`Remove ${col.label}`}
                        title="Remove"
                        onClick={() => onChange(removeColumn(keys, key))}
                      >
                        <Cross />
                      </button>
                    </span>
                  )}
                </li>
              );
            })}
          </ol>
        </div>

        <div className={styles.fieldset}>
          <h3 className={styles.fieldLabel}>Add a column</h3>
          {hidden.length === 0 ? (
            <p className={styles.muted}>Every column is already shown.</p>
          ) : (
            <ul className={styles.addList}>
              {hidden.map((c) => (
                <li key={c.key}>
                  <button type="button" className={styles.addBtn} onClick={() => onChange(addColumn(keys, c.key))}>
                    <Plus />
                    {c.label}
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className={styles.dialogFoot}>
          <button
            type="button"
            className={styles.secondaryBtn}
            disabled={isDefault(keys)}
            onClick={() => onChange(DEFAULT_COLUMNS.slice())}
          >
            Reset to default
          </button>
          <button type="button" className={styles.primaryBtn} onClick={onClose}>
            Done
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}
