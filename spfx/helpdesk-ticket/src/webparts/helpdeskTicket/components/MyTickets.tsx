// "My tickets": the tickets this person raised, split into open and closed.
import * as React from 'react';
import { TicketsApi } from '../sharepoint/ticketsApi';
import { matchesSearch, splitOpenClosed, TicketRow } from '../tickets/myTickets';
import { useInstanceId, useMountedRef } from './hooks';
import { RefreshIcon, SearchIcon } from './icons';
import { friendlyUpdated } from './ticketFormat';
import styles from './HelpdeskTicket.module.scss';

export interface MyTicketsProps {
  api: TicketsApi;
  userId: number;
  /** True while this screen is the one on show; it loads each time it comes into view. */
  active: boolean;
  onRaise: () => void;
}

type Tab = 'open' | 'closed';

const STATUS_COLOUR: Record<string, string> = {
  New: '#8a8886',
  'In progress': '#0b5cad',
  'On Hold': '#ca5010',
  Completed: '#107c10',
  Cancelled: '#8a8886'
};

const messageOf = (e: unknown): string => (e instanceof Error ? e.message : 'Something went wrong.');

export function MyTickets(props: MyTicketsProps): React.ReactElement {
  const { api, userId, active, onRaise } = props;
  const [rows, setRows] = React.useState<TicketRow[] | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [tab, setTab] = React.useState<Tab>('open');
  const [query, setQuery] = React.useState('');
  const seq = React.useRef(0);
  const mounted = useMountedRef();
  const tabRefs = { open: React.useRef<HTMLButtonElement>(null), closed: React.useRef<HTMLButtonElement>(null) };
  const ids = {
    title: useInstanceId('mt-title'),
    search: useInstanceId('mt-search'),
    panel: useInstanceId('mt-panel')
  };

  const load = React.useCallback(async (): Promise<void> => {
    const mine = ++seq.current;
    setLoading(true);
    setError('');
    try {
      const result = await api.myTickets(userId);
      if (mounted.current && mine === seq.current) setRows(result);
    } catch (e) {
      if (mounted.current && mine === seq.current) setError(messageOf(e));
    } finally {
      if (mounted.current && mine === seq.current) setLoading(false);
    }
  }, [api, userId, mounted]);

  React.useEffect(() => {
    if (active) load().catch(() => undefined);
  }, [active, load]);

  const term = query.trim();
  const filtered = rows ? rows.filter((r) => matchesSearch(r, query)) : [];
  const { open, closed } = splitOpenClosed(filtered);
  const shown = tab === 'open' ? open : closed;
  const now = new Date();

  const onTabKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    let next: Tab | undefined;
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') next = tab === 'open' ? 'closed' : 'open';
    else if (e.key === 'Home') next = 'open';
    else if (e.key === 'End') next = 'closed';
    if (!next) return;
    e.preventDefault();
    setTab(next);
    tabRefs[next].current?.focus();
  };

  const showTabs = !!rows && rows.length > 0;
  const emptyAll = !!rows && rows.length === 0 && !loading && !error;

  return (
    <section className={styles.card} aria-labelledby={ids.title}>
      <div className={styles.listHead}>
        <div className={styles.listTitleBox}>
          <h1 id={ids.title} className={styles.pageTitle}>
            My tickets
          </h1>
          <p className={styles.muted}>Tickets you raised. Only you and IT can see these.</p>
        </div>
        <div className={styles.listActions}>
          <div className={styles.searchBox}>
            <SearchIcon className={styles.searchIcon} />
            <label htmlFor={ids.search} className={styles.visuallyHidden}>
              Search my tickets
            </label>
            <input
              id={ids.search}
              type="search"
              className={styles.searchInput}
              placeholder="Search number or description"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <button
            type="button"
            className={styles.iconBtn}
            aria-label="Refresh"
            title="Refresh"
            disabled={loading}
            onClick={() => load()}
          >
            <RefreshIcon className={loading ? styles.spinIcon : undefined} />
          </button>
          <button type="button" className={styles.primaryBtn} onClick={onRaise}>
            Raise a ticket
          </button>
        </div>
      </div>

      {error && (
        <div className={styles.errorBox} role="alert">
          <span>{error}</span>
          <button type="button" className={styles.secondaryBtn} onClick={() => load()}>
            Try again
          </button>
        </div>
      )}

      {loading && rows === null && !error && (
        <div className={styles.stateBox} role="status">
          <span className={styles.spinner} aria-hidden="true" />
          Loading your tickets…
        </div>
      )}

      {emptyAll && (
        <div className={styles.stateBox}>
          <span>You haven&apos;t raised any tickets yet.</span>
          <button type="button" className={styles.primaryBtn} onClick={onRaise}>
            Raise a ticket
          </button>
        </div>
      )}

      {showTabs && (
        <>
          <div role="tablist" aria-label="Which tickets" className={styles.tabs} onKeyDown={onTabKey}>
            {(['open', 'closed'] as Tab[]).map((t) => {
              const selected = tab === t;
              const count = t === 'open' ? open.length : closed.length;
              return (
                <button
                  key={t}
                  ref={tabRefs[t]}
                  type="button"
                  role="tab"
                  id={`${ids.panel}-${t}-tab`}
                  aria-selected={selected}
                  aria-controls={`${ids.panel}-panel`}
                  tabIndex={selected ? 0 : -1}
                  className={selected ? styles.tabOn : styles.tab}
                  onClick={() => setTab(t)}
                >
                  {t === 'open' ? 'Open' : 'Closed'} ({count})
                </button>
              );
            })}
          </div>

          <div
            role="tabpanel"
            id={`${ids.panel}-panel`}
            aria-labelledby={`${ids.panel}-${tab}-tab`}
            className={styles.tableWrap}
          >
            {term && filtered.length === 0 ? (
              <div className={styles.stateBox} role="status">
                No tickets match “{term}”.
              </div>
            ) : shown.length === 0 ? (
              <div className={styles.stateBox} role="status">
                {tab === 'open' ? 'No open tickets.' : 'No closed tickets.'}
              </div>
            ) : (
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th scope="col">Number</th>
                    <th scope="col">Short description</th>
                    <th scope="col">Category</th>
                    <th scope="col">Priority</th>
                    <th scope="col">Status</th>
                    <th scope="col">Updated</th>
                  </tr>
                </thead>
                <tbody>
                  {shown.map((r) => (
                    <tr key={r.id}>
                      <td>
                        <a href={api.displayFormUrl(r.id)} className={r.number ? undefined : styles.pending}>
                          {/* The flow numbers a ticket within moments; one it never numbered is not "pending". */}
                          {r.number || (now.getTime() - new Date(r.created).getTime() < 15 * 60 * 1000 ? 'Pending' : 'No number')}
                        </a>
                      </td>
                      <td>{r.title}</td>
                      <td className={styles.muted}>{r.category}</td>
                      <td>{r.priority}</td>
                      <td>
                        <span className={styles.status}>
                          <span
                            className={styles.dot}
                            style={{ background: STATUS_COLOUR[r.status] ?? '#8a8886' }}
                            aria-hidden="true"
                          />
                          {r.status}
                        </span>
                      </td>
                      <td className={styles.muted}>{friendlyUpdated(r.modified, now)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}
    </section>
  );
}
