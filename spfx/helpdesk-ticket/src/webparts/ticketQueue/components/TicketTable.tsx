// The ticket table for the current page, with sortable headers and a pager.
import * as React from 'react';
import { QueueApi } from '../sharepoint/queueApi';
import { SortKey } from '../queue/queueFilter';
import { QueueTicket } from '../queue/ticket';
import { PAGE_SIZE, formatWhen, raisedLine } from './queueFormat';
import { SortIcon } from './queueIcons';
import styles from './TicketQueue.module.scss';

export interface SortState {
  key: SortKey;
  desc: boolean;
}

export interface TicketTableProps {
  api: QueueApi;
  /** Every matching ticket, already sorted. The table shows one page of them. */
  tickets: QueueTicket[];
  /** The page on show, already clamped by the screen. */
  page: number;
  sort: SortState;
  onSort(key: SortKey): void;
  onPage(page: number): void;
}

interface Column {
  label: string;
  sort?: SortKey;
}

const COLUMNS: Column[] = [
  { label: 'Ticket ID', sort: 'number' },
  { label: 'Short description' },
  { label: 'Category' },
  { label: 'Priority', sort: 'priority' },
  { label: 'Status', sort: 'status' },
  { label: 'Assigned to', sort: 'assignedTo' },
  { label: 'Raised on', sort: 'created' },
  { label: 'Updated', sort: 'modified' }
];

const STATUS_COLOUR: Record<string, string> = {
  New: '#8a8886',
  'In progress': '#0b5cad',
  'On Hold': '#ca5010',
  Completed: '#107c10',
  Cancelled: '#8a8886'
};

const PRIORITY_CLASS: Record<string, string> = {
  Critical: styles.critical,
  High: styles.high
};

export function TicketTable(props: TicketTableProps): React.ReactElement {
  const { api, tickets, page, sort, onSort, onPage } = props;
  const pages = Math.max(1, Math.ceil(tickets.length / PAGE_SIZE));
  const from = page * PAGE_SIZE;
  const shown = tickets.slice(from, from + PAGE_SIZE);
  const to = Math.min(tickets.length, from + PAGE_SIZE);

  return (
    <>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <caption className={styles.visuallyHidden}>Tickets, {tickets.length} in total</caption>
          <thead>
            <tr>
              {COLUMNS.map((c) => {
                if (!c.sort) {
                  return (
                    <th key={c.label} scope="col">
                      {c.label}
                    </th>
                  );
                }
                const active = sort.key === c.sort;
                const ariaSort: 'ascending' | 'descending' | 'none' = active
                  ? sort.desc
                    ? 'descending'
                    : 'ascending'
                  : 'none';
                return (
                  <th key={c.label} scope="col" aria-sort={ariaSort}>
                    <button type="button" className={styles.sortBtn} onClick={() => onSort(c.sort as SortKey)}>
                      {c.label}
                      <SortIcon direction={active ? (sort.desc ? 'desc' : 'asc') : undefined} />
                    </button>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {shown.map((t) => {
              const raised = raisedLine(t);
              return (
                <tr key={t.id}>
                  <td className={styles.nowrap}>
                    <a
                      href={api.displayFormUrl(t.id)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={t.number ? styles.link : styles.linkMuted}
                    >
                      {t.number || 'No number'}
                    </a>
                  </td>
                  <td className={styles.titleCell}>
                    <span className={styles.titleText}>{t.title || '(no title)'}</span>
                    {raised && <span className={styles.subLine}>{raised}</span>}
                  </td>
                  <td className={styles.muted}>{t.category}</td>
                  <td className={`${styles.nowrap} ${PRIORITY_CLASS[t.priority] ?? ''}`}>{t.priority}</td>
                  <td className={styles.nowrap}>
                    <span className={styles.status}>
                      <span
                        className={styles.dot}
                        style={{ background: STATUS_COLOUR[t.status] ?? '#8a8886' }}
                        aria-hidden="true"
                      />
                      {t.status}
                    </span>
                  </td>
                  <td className={styles.nowrap}>
                    {t.assignedTo ? t.assignedTo.name : <span className={styles.unassigned}>Unassigned</span>}
                  </td>
                  <td className={`${styles.nowrap} ${styles.muted}`}>
                    <time dateTime={t.created}>{formatWhen(t.created)}</time>
                  </td>
                  <td className={`${styles.nowrap} ${styles.muted}`}>
                    <time dateTime={t.modified}>{formatWhen(t.modified)}</time>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <nav className={styles.pager} aria-label="Pages">
        <button type="button" className={styles.secondaryBtn} disabled={page === 0} onClick={() => onPage(page - 1)}>
          Previous
        </button>
        <span className={styles.muted} role="status">
          {`${from + 1}–${to} of ${tickets.length}`}
        </span>
        <button
          type="button"
          className={styles.secondaryBtn}
          disabled={page >= pages - 1}
          onClick={() => onPage(page + 1)}
        >
          Next
        </button>
      </nav>
    </>
  );
}
