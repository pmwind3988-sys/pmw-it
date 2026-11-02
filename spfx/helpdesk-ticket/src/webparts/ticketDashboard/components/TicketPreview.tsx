// The panel a tap on a chart opens: the tickets behind that mark, a few at a
// time, each with its subject, the start of its description and a link to the
// ticket. A drawer on the right on a wide screen, a sheet from the bottom on a
// phone (Dashboard.module.scss decides which).
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { QueueTicket } from '../../ticketQueue/queue/ticket';
import { useInstanceId } from '../../helpdeskTicket/components/hooks';
import { Drill, DrillList } from '../stats/drill';
import { isOpen } from '../stats/stats';
import { ticketsText } from './chartMath';
import styles from './Dashboard.module.scss';

const FIRST_PAGE = 5;
const PAGE = 10;

export interface TicketPreviewProps {
  list: DrillList;
  /** The ticket's own page, opened in a new tab. */
  urlOf: (id: number) => string;
  /** A tab inside the panel (Raised / Closed) picks the other half of the same point. */
  onPick: (d: Drill) => void;
  /** Which tab is showing, by its label's position. */
  tabIndex?: number;
  onClose: () => void;
}

const STATUS_CLASS: Record<string, string> = {
  New: styles.pillNew,
  'In progress': styles.pillProgress,
  'On Hold': styles.pillHold
};

function dayText(iso: string, now: Date): string {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const opts: Intl.DateTimeFormatOptions =
    d.getFullYear() === now.getFullYear() ? { day: 'numeric', month: 'short' } : { day: 'numeric', month: 'short', year: 'numeric' };
  return d.toLocaleDateString('en-GB', opts);
}

function whereText(t: QueueTicket, now: Date): string {
  if (t.status === 'Completed') {
    const at = t.completedDate || t.modified;
    return `Closed ${dayText(at, now)}`;
  }
  if (!isOpen(t)) return 'Cancelled';
  return t.assignedTo ? `With ${t.assignedTo.name}` : 'Nobody has it yet';
}

function TicketRow(props: { t: QueueTicket; url: string; now: Date }): React.ReactElement {
  const { t, url, now } = props;
  const urgent = t.priority === 'Critical' || t.priority === 'High';
  const meta = [`Raised ${dayText(t.created, now)}`, t.department, whereText(t, now)].filter((x) => x !== '').join(' · ');
  return (
    <li className={styles.pvItem}>
      <div className={styles.pvTags}>
        {t.number !== '' && <span className={styles.pvNumber}>{t.number}</span>}
        <span className={`${styles.pill} ${STATUS_CLASS[t.status] ?? ''}`}>{t.status === 'On Hold' ? 'On hold' : t.status}</span>
        {urgent && <span className={`${styles.pill} ${styles.pillUrgent}`}>{t.priority}</span>}
      </div>
      <span className={styles.pvTitle}>{t.title || 'No subject'}</span>
      <span className={styles.pvDesc}>{t.description || 'No description.'}</span>
      <div className={styles.pvFoot}>
        <span className={styles.pvMeta}>{meta}</span>
        <a className={styles.pvLink} href={url} target="_blank" rel="noopener noreferrer">
          Open ticket
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
            <path d="M14 4h6v6" />
            <path d="M20 4 10 14" />
            <path d="M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5" />
          </svg>
          <span className={styles.visuallyHidden}>(opens in a new tab)</span>
        </a>
      </div>
    </li>
  );
}

export function TicketPreview(props: TicketPreviewProps): React.ReactElement {
  const { list, urlOf, onPick, tabIndex, onClose } = props;
  const [limit, setLimit] = React.useState(FIRST_PAGE);
  const panel = React.useRef<HTMLDivElement>(null);
  const titleId = useInstanceId('pv-title');
  const now = new Date();

  // Focus moves into the panel, and back to whatever opened it when it closes.
  React.useEffect(() => {
    const opener = document.activeElement as HTMLElement | null;
    panel.current?.focus();
    return () => {
      if (opener && document.contains(opener)) opener.focus();
    };
  }, []);

  const onKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key !== 'Tab' || !panel.current) return;
    // Keep Tab inside the panel while it is open.
    const items = Array.from(panel.current.querySelectorAll<HTMLElement>('button:not([disabled]), a[href]'));
    if (!items.length) return;
    const first = items[0];
    const last = items[items.length - 1];
    const active = document.activeElement;
    if (e.shiftKey && (active === first || active === panel.current)) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && active === last) {
      e.preventDefault();
      first.focus();
    }
  };

  const total = list.tickets.length;
  const shown = list.tickets.slice(0, limit);

  return ReactDom.createPortal(
    <div className={styles.pvScrim} onClick={onClose}>
      <div
        ref={panel}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        tabIndex={-1}
        className={styles.pvPanel}
        onKeyDown={onKey}
        onClick={(e) => e.stopPropagation()}
      >
        <span className={styles.pvGrip} aria-hidden="true" />
        <div className={styles.pvHead}>
          <div className={styles.pvHeadText}>
            <span className={styles.pvKicker}>{list.kicker}</span>
            <h2 id={titleId} className={styles.pvHeading}>
              {list.title}
            </h2>
            <span className={styles.pvCount}>{`${ticketsText(total)} · ${list.order}`}</span>
          </div>
          <button type="button" className={styles.pvClose} aria-label="Close" onClick={onClose}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" aria-hidden="true" focusable="false">
              <path d="M6 6l12 12" />
              <path d="M18 6 6 18" />
            </svg>
          </button>
        </div>

        {list.tabs !== undefined && (
          <div className={styles.pvTabs} role="group" aria-label="Which tickets">
            {list.tabs.map((tab, i) => (
              <button
                key={tab.label}
                type="button"
                className={styles.toggle}
                aria-pressed={i === (tabIndex ?? 0)}
                onClick={() => {
                  setLimit(FIRST_PAGE);
                  onPick(tab.drill);
                }}
              >
                {tab.label}
              </button>
            ))}
          </div>
        )}

        {total === 0 ? (
          <p className={styles.pvEmpty}>No tickets here.</p>
        ) : (
          <ul className={styles.pvList}>
            {shown.map((t) => (
              <TicketRow key={t.id} t={t} url={urlOf(t.id)} now={now} />
            ))}
          </ul>
        )}

        {total > 0 && (
          <div className={styles.pvMore}>
            <span className={styles.pvCount} role="status">{`Showing ${shown.length} of ${total}`}</span>
            {total > limit && (
              <button type="button" className={styles.secondaryBtn} onClick={() => setLimit(limit + PAGE)}>
                {`Show ${Math.min(PAGE, total - limit)} more`}
              </button>
            )}
          </div>
        )}
      </div>
    </div>,
    document.body
  );
}
