// The IT ticket queue. Owners only: anyone else gets nothing, and no tickets
// are loaded for them. The owner's screen: find tickets, filter them, sort,
// page through them, and export what the filters show.
import * as React from 'react';
import { ITicketQueueProps } from './ITicketQueueProps';
import { applyFilter, assigneesOf, defaultFilter, QueueFilter, SortKey, sortTickets } from '../queue/queueFilter';
import { describe as describeRange } from '../queue/dateRange';
import { PersonRef, QueueTicket } from '../queue/ticket';
import { useInstanceId, useMountedRef } from '../../helpdeskTicket/components/hooks';
import { RefreshIcon } from '../../helpdeskTicket/components/icons';
import {
  assigneeNames,
  assigneeSummary,
  countText,
  ExportContext,
  initialDesc,
  pageCount,
  parseFilter,
  sortDescription
} from './queueFormat';
import { FILTER_KEY, readStored, writeStored } from './queueStorage';
import { QueueFilters } from './QueueFilters';
import { SortState, TicketTable } from './TicketTable';
import { ExportDialog } from './ExportDialog';
import styles from './TicketQueue.module.scss';

const messageOf = (e: unknown): string => (e instanceof Error ? e.message : 'Something went wrong.');

export default function TicketQueue(props: ITicketQueueProps): React.ReactElement {
  if (!props.isOwner) {
    return props.isEditing ? (
      <p className={styles.note}>IT ticket queue — only owners of this site can see and use it.</p>
    ) : (
      <></>
    );
  }
  return <QueueScreen api={props.api} me={props.me} />;
}

interface QueueScreenProps {
  api: ITicketQueueProps['api'];
  me: PersonRef;
}

function QueueScreen(props: QueueScreenProps): React.ReactElement {
  const { api, me } = props;
  const [rows, setRows] = React.useState<QueueTicket[] | undefined>(undefined);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [filter, setFilter] = React.useState<QueueFilter>(() => parseFilter(readStored(FILTER_KEY), me.id));
  const [searchText, setSearchText] = React.useState<string>(filter.search);
  const [sort, setSort] = React.useState<SortState>({ key: 'created', desc: true });
  const [page, setPage] = React.useState(0);
  const [exportOpen, setExportOpen] = React.useState(false);
  const seq = React.useRef(0);
  const mounted = useMountedRef();
  const exportBtn = React.useRef<HTMLButtonElement>(null);
  const ids = {
    title: useInstanceId('tq-title'),
    list: useInstanceId('tq-list')
  };

  const load = React.useCallback(async (): Promise<void> => {
    const mine = ++seq.current;
    setLoading(true);
    setError('');
    try {
      const result = await api.allTickets();
      if (mounted.current && mine === seq.current) setRows(result);
    } catch (e) {
      if (mounted.current && mine === seq.current) setError(messageOf(e));
    } finally {
      if (mounted.current && mine === seq.current) setLoading(false);
    }
  }, [api, mounted]);

  React.useEffect(() => {
    load().catch(() => undefined);
  }, [load]);

  React.useEffect(() => {
    writeStored(FILTER_KEY, JSON.stringify(filter));
  }, [filter]);

  // The search box applies 200ms after the last keystroke, not on every one.
  React.useEffect(() => {
    const timer = window.setTimeout(() => {
      setFilter((f) => (f.search === searchText ? f : { ...f, search: searchText }));
    }, 200);
    return () => window.clearTimeout(timer);
  }, [searchText]);

  const all = rows ?? [];
  const now = new Date();
  const nameById = new Map<number, string>(
    assigneesOf(all).map((p): [number, string] => [p.person.id, p.person.name])
  );
  const nameOf = (id: number): string => nameById.get(id) ?? 'Someone else';

  const matched = applyFilter(all, filter, now);
  const sorted = sortTickets(matched, sort.key, sort.desc);
  // The page is clamped while drawing: a filter that shrinks the list must not leave an empty page up.
  const current = Math.min(page, pageCount(sorted.length) - 1);

  const assigneeText = assigneeSummary(filter.assignees, me, nameOf);
  const dateText = describeRange(filter.date, now);
  const names = filter.assignees.length ? assigneeNames(filter.assignees, me, nameOf, me.name || 'Me') : [];
  const resultText = [countText(matched.length), dateText, names.length ? `Assigned to ${names.join(', ')}` : '']
    .filter((s) => s !== '')
    .join(' · ');

  const patch = (change: Partial<QueueFilter>): void => setFilter((f) => ({ ...f, ...change }));
  const reset = (): void => {
    setFilter(defaultFilter(me.id));
    setSearchText('');
    setPage(0);
  };

  const onSort = (key: SortKey): void => {
    setSort((s) => (s.key === key ? { key, desc: !s.desc } : { key, desc: initialDesc(key) }));
  };

  const exportContext: ExportContext = {
    exportedBy: me.name,
    assignees: assigneeText,
    raised: dateText,
    statuses: filter.statuses,
    search: filter.search,
    sortText: sortDescription(sort.key, sort.desc),
    fileParts: [assigneeText, dateText]
  };

  const closeExport = (): void => {
    setExportOpen(false);
    exportBtn.current?.focus();
  };

  return (
    <div className={styles.page}>
      <section className={styles.card} aria-labelledby={ids.title}>
        <div className={styles.head}>
          <div className={styles.headText}>
            <h1 id={ids.title} className={styles.title}>
              Ticket queue
            </h1>
            <p className={styles.muted}>Only owners of this site see this page.</p>
          </div>
          <div className={styles.headActions}>
            <button
              type="button"
              className={styles.iconBtn}
              aria-label="Refresh"
              title="Refresh"
              disabled={loading}
              onClick={() => {
                load().catch(() => undefined);
              }}
            >
              <RefreshIcon className={loading ? styles.spin : undefined} />
            </button>
            <button
              ref={exportBtn}
              type="button"
              className={styles.primaryBtn}
              disabled={matched.length === 0}
              onClick={() => setExportOpen(true)}
            >
              Export
            </button>
          </div>
        </div>
      </section>

      <section className={styles.card} aria-label="Filters">
        <QueueFilters
          filter={filter}
          me={me}
          rows={all}
          assigneeText={assigneeText}
          searchText={searchText}
          onSearchText={setSearchText}
          onChange={patch}
          onReset={reset}
        />
      </section>

      <section className={styles.card} aria-labelledby={ids.list}>
        <h2 id={ids.list} className={styles.visuallyHidden}>
          Tickets
        </h2>
        <p className={styles.resultLine} role="status">
          {rows !== undefined && resultText}
        </p>

        {error && (
          <div className={styles.errorBox} role="alert">
            <span>{error}</span>
            <button
              type="button"
              className={styles.secondaryBtn}
              onClick={() => {
                load().catch(() => undefined);
              }}
            >
              Try again
            </button>
          </div>
        )}

        {loading && rows === undefined && !error && (
          <div className={styles.stateBox} role="status">
            <span className={styles.spinner} aria-hidden="true" />
            Loading tickets…
          </div>
        )}

        {rows !== undefined && rows.length === 0 && !error && (
          <div className={styles.stateBox}>No tickets in the list yet.</div>
        )}

        {rows !== undefined && rows.length > 0 && matched.length === 0 && (
          <div className={styles.stateBox} role="status">
            <span>No tickets match these filters.</span>
            <button type="button" className={styles.secondaryBtn} onClick={reset}>
              Reset filters
            </button>
          </div>
        )}

        {matched.length > 0 && (
          <TicketTable
            api={api}
            tickets={sorted}
            page={current}
            sort={sort}
            onSort={onSort}
            onPage={setPage}
          />
        )}
      </section>

      {exportOpen && <ExportDialog rows={sorted} context={exportContext} onClose={closeExport} />}
    </div>
  );
}
