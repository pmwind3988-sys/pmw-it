// The ticket table for the current page, with sortable headers and a pager.
import * as React from 'react';
import { QueueApi } from '../sharepoint/queueApi';
import { SortKey } from '../queue/queueFilter';
import { columnOf, TableColumn } from '../queue/tableColumns';
import { QueueTicket } from '../queue/ticket';
import { timeToCompleteText } from '../queue/timeToComplete';
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
  /** The columns chosen, in order (tableColumns.ts). */
  columns: string[];
}

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

const people = (list: { name: string }[]): string => list.map((p) => p.name).join('; ');

/** One cell of one column. Long text is clamped to a few lines; hover shows it all. */
function Cell(props: { column: string; t: QueueTicket; api: QueueApi }): React.ReactElement {
  const { column, t, api } = props;
  const plain = (value: string | number | undefined, className = styles.muted): React.ReactElement => (
    <td className={`${styles.nowrap} ${className}`}>{value === undefined ? '' : value}</td>
  );
  const long = (value: string): React.ReactElement => (
    <td className={styles.longCell}>
      <span className={styles.clamp} title={value}>
        {value}
      </span>
    </td>
  );
  const when = (iso: string): React.ReactElement => (
    <td className={`${styles.nowrap} ${styles.muted}`}>{iso && <time dateTime={iso}>{formatWhen(iso)}</time>}</td>
  );
  switch (column) {
    case 'number':
      return (
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
      );
    case 'title': {
      const raised = raisedLine(t);
      return (
        <td className={styles.titleCell}>
          <span className={styles.titleText}>{t.title || '(no title)'}</span>
          {raised && <span className={styles.subLine}>{raised}</span>}
        </td>
      );
    }
    case 'category':
      return <td className={styles.muted}>{t.category}</td>;
    case 'priority':
      return <td className={`${styles.nowrap} ${PRIORITY_CLASS[t.priority] ?? ''}`}>{t.priority}</td>;
    case 'status':
      return (
        <td className={styles.nowrap}>
          <span className={styles.status}>
            <span className={styles.dot} style={{ background: STATUS_COLOUR[t.status] ?? '#8a8886' }} aria-hidden="true" />
            {t.status}
          </span>
        </td>
      );
    case 'assignedTo':
      return (
        <td className={styles.nowrap}>
          {t.assignedTo ? t.assignedTo.name : <span className={styles.unassigned}>Unassigned</span>}
        </td>
      );
    case 'created':
      return when(t.created);
    case 'modified':
      return when(t.modified);
    case 'timeToComplete':
      return plain(timeToCompleteText(t), styles.durationCell);
    case 'completedDate':
      return when(t.status === 'Completed' ? t.completedDate || t.modified : '');
    case 'completedBy':
      return plain(t.completedBy?.name);
    case 'raisedBy':
      return plain(t.raisedBy?.name);
    case 'raisedByEmail':
      return plain(t.raisedBy?.email);
    case 'department':
      return plain(t.department);
    case 'description':
      return long(t.description);
    case 'cc':
      return long(people(t.cc));
    case 'rootCause':
      return plain(t.rootCause);
    case 'resolution':
      return long(t.resolution);
    case 'comments':
      return long(t.comments);
    case 'holdCount':
      return plain(t.holdCount);
    case 'totalHoldDays':
      return plain(t.totalHoldDays);
    case 'kbCandidate':
      return plain(t.kbCandidate ? 'Yes' : 'No');
    default:
      return <td />;
  }
}

export function TicketTable(props: TicketTableProps): React.ReactElement {
  const { api, tickets, page, sort, onSort, onPage } = props;
  const columns = props.columns.map(columnOf).filter((c): c is TableColumn => c !== undefined);
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
              {columns.map((c) => {
                if (!c.sort) {
                  return (
                    <th key={c.key} scope="col">
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
                  <th key={c.key} scope="col" aria-sort={ariaSort}>
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
            {shown.map((t) => (
              <tr key={t.id}>
                {columns.map((c) => (
                  <Cell key={c.key} column={c.key} t={t} api={api} />
                ))}
              </tr>
            ))}
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
