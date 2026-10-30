// The IT helpdesk dashboard. Owners only: anyone else gets nothing, and no
// tickets are loaded for them. The figures, charts and tables are all worked out
// in stats/stats.ts; this file lays them out and wires the queue's filters and
// export to them.
import * as React from 'react';
import { ITicketDashboardProps } from './ITicketDashboardProps';
import { applyFilter, assigneesOf, QueueFilter, sortTickets } from '../../ticketQueue/queue/queueFilter';
import { describe as describeRange, Range, rangeOf } from '../../ticketQueue/queue/dateRange';
import { PersonRef, QueueTicket } from '../../ticketQueue/queue/ticket';
import { QueueApi } from '../../ticketQueue/sharepoint/queueApi';
import { useInstanceId, useMountedRef } from '../../helpdeskTicket/components/hooks';
import { RefreshIcon } from '../../helpdeskTicket/components/icons';
import {
  assigneeSummary,
  countText,
  ExportContext,
  parseFilter,
  sortDescription
} from '../../ticketQueue/components/queueFormat';
import { readStored, writeStored } from '../../ticketQueue/components/queueStorage';
import { QueueFilters } from '../../ticketQueue/components/QueueFilters';
import { ExportDialog } from '../../ticketQueue/components/ExportDialog';
import { DownloadIcon } from '../../ticketQueue/components/queueIcons';
import {
  ages,
  arrivals,
  byArea,
  byPriority,
  changeText,
  days1,
  gaps,
  headline,
  holders,
  leaderboard,
  Period,
  periodOf,
  speeds,
  stories,
  trend
} from '../stats/stats';
import { ChartCard, KpiTile, LegendItem, StoryCard } from './DashboardParts';
import { TrendPlot, TrendTable } from './TrendPlot';
import {
  AREA_LEGEND,
  AreaBars,
  AreaTable,
  HOLDER_LEGEND,
  HoldersBars,
  HoldersTable,
  Leaderboard,
  LeaderboardTable,
  PriorityBar,
  PriorityTable,
  priorityRows
} from './Breakdowns';
import { AgeBars, AgeTable, GapBars, GapTable, HeatLegend, HeatMap, HeatTable, SpeedBars, SpeedTable } from './Timing';
import { MONTH_NAMES, PALETTE, percentText } from './chartMath';
import styles from './Dashboard.module.scss';

/** Kept apart from the queue's remembered filter, so the two screens never overwrite each other. */
export const FILTER_KEY = 'ticketDashboard:filter:v1';

/** Everyone, any time, every status. Reset filters returns here. */
const everyone = (): QueueFilter => ({ assignees: [], date: { preset: 'all' }, statuses: [], search: '' });

const messageOf = (e: unknown): string => (e instanceof Error ? e.message : 'Something went wrong.');

function isStoredObject(raw: string): boolean {
  try {
    const parsed: unknown = JSON.parse(raw);
    return typeof parsed === 'object' && !!parsed && !Array.isArray(parsed);
  } catch {
    return false;
  }
}

/** The remembered filter, or everyone-any-time when nothing usable is remembered. */
function startFilter(raw: string | undefined, myId: number): QueueFilter {
  if (raw === undefined || !isStoredObject(raw)) return everyone();
  return parseFilter(raw, myId);
}

export default function TicketDashboard(props: ITicketDashboardProps): React.ReactElement {
  if (!props.isOwner) {
    return props.isEditing ? (
      <p className={styles.note}>IT helpdesk dashboard — only owners of this site can see and use it.</p>
    ) : (
      <></>
    );
  }
  return <DashboardScreen api={props.api} me={props.me} />;
}

interface DashboardScreenProps {
  api: QueueApi;
  me: PersonRef;
}

function DashboardScreen(props: DashboardScreenProps): React.ReactElement {
  const { api, me } = props;
  const [rows, setRows] = React.useState<QueueTicket[] | undefined>(undefined);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [filter, setFilter] = React.useState<QueueFilter>(() => startFilter(readStored(FILTER_KEY), me.id));
  const [searchText, setSearchText] = React.useState<string>(filter.search);
  const [exportOpen, setExportOpen] = React.useState(false);
  const seq = React.useRef(0);
  const mounted = useMountedRef();
  const exportBtn = React.useRef<HTMLButtonElement>(null);
  const ids = {
    title: useInstanceId('td-title'),
    list: useInstanceId('td-list')
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

  // "Present" figures ignore the dates; everything the filters pick follows them.
  const scope = applyFilter(all, { ...filter, date: { preset: 'all' } }, now);
  const inView = applyFilter(all, filter, now);
  const range: Range = rangeOf(filter.date, now);
  const period = periodOf(filter.date, now);

  const assigneeText = assigneeSummary(filter.assignees, me, nameOf);
  const dateText = describeRange(filter.date, now);
  const resultText = [countText(inView.length), dateText, assigneeText].join(' · ');

  const patch = (change: Partial<QueueFilter>): void => setFilter((f) => ({ ...f, ...change }));
  const reset = (): void => {
    setFilter(everyone());
    setSearchText('');
  };

  const exportContext: ExportContext = {
    exportedBy: me.name,
    assignees: assigneeText,
    raised: dateText,
    statuses: filter.statuses,
    search: filter.search,
    sortText: sortDescription('created', true),
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
              IT helpdesk dashboard
            </h1>
            <p className={styles.muted}>
              How the helpdesk is doing, for the tickets the filters below pick. Only owners of this site see this page.
            </p>
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
              disabled={inView.length === 0}
              onClick={() => setExportOpen(true)}
            >
              <DownloadIcon />
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
        <div className={styles.resultBox}>
          <p className={styles.resultLine} role="status">
            {rows !== undefined && resultText}
          </p>
          <p className={styles.hint}>Every chart and figure below follows these filters, and so does Export.</p>
          <p className={styles.hint}>Who holds what and the month&apos;s leaderboard always show the present.</p>
        </div>
      </section>

      <section className={styles.bodyCard} aria-labelledby={ids.list}>
        <h2 id={ids.list} className={styles.visuallyHidden}>
          Figures
        </h2>

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

        {rows !== undefined && rows.length > 0 && (
          <DashboardBody api={api} scope={scope} inView={inView} range={range} period={period} now={now} />
        )}
      </section>

      {exportOpen && (
        <ExportDialog rows={sortTickets(inView, 'created', true)} context={exportContext} onClose={closeExport} />
      )}
    </div>
  );
}

interface DashboardBodyProps {
  api: QueueApi;
  /** Tickets the filters' people and statuses pick, any date: the present. */
  scope: QueueTicket[];
  /** Tickets the whole filter picks, dates included. */
  inView: QueueTicket[];
  range: Range;
  /** What the headline cards count: the chosen dates, or this month for "Any time". */
  period: Period;
  now: Date;
}

const AlertIcon = (): React.ReactElement => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
    <path d="M12 9v4" />
    <path d="M12 17h.01" />
    <path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z" />
  </svg>
);

function DashboardBody(props: DashboardBodyProps): React.ReactElement {
  const { api, scope, inView, range, period, now } = props;
  const empty = inView.length === 0;

  const t = trend(scope, range, now);
  const h = headline(scope, inView, period, now);
  const heat = arrivals(inView);
  const team = holders(scope);
  const board = leaderboard(scope, now);
  const area = byArea(inView);
  const pri = priorityRows(byPriority(inView));
  const sp = speeds(inView);
  const ag = ages(scope, now);
  const gp = gaps(inView);
  const story = stories(h, heat, period);
  const monthName = MONTH_NAMES[now.getMonth()];

  const trendTitle =
    t.grain === 'month'
      ? 'Raised vs closed, month by month'
      : t.grain === 'week'
        ? 'Raised vs closed, week by week'
        : 'Raised vs closed, day by day';
  const openTitle = `Still open at the end of each ${t.grain}`;

  const oldest = h.oldestOpen;
  const oldBucket = ag[ag.length - 1];
  const criticalNote = h.criticalInView > 0 ? `1 in ${Math.round(inView.length / h.criticalInView)} tickets is Critical.` : undefined;

  const kpis = [
    {
      label: 'Open now',
      value: String(h.openNow),
      note: h.unassignedOpen > 0 ? `${h.unassignedOpen} nobody has picked up` : 'All picked up',
      noteColour: h.unassignedOpen > 0 ? PALETTE.critical : '#107c10'
    },
    {
      label: `Raised ${period.label}`,
      value: String(h.raised),
      note: changeText(h.raised, h.raisedBefore, period.prevLabel)
    },
    {
      label: `Closed ${period.label}`,
      value: String(h.closed),
      note: changeText(h.closed, h.closedBefore, period.prevLabel),
      noteColour: h.closed >= h.raised ? '#107c10' : undefined
    },
    {
      label: 'Median time to close',
      value: h.medianDays !== undefined ? days1(h.medianDays) : '—',
      note: 'for these tickets'
    },
    {
      label: 'Oldest open ticket',
      value: oldest !== undefined ? `${oldest.days} d` : '—',
      note: oldest !== undefined ? oldest.ticket.number : 'No open tickets',
      href: oldest !== undefined ? api.displayFormUrl(oldest.ticket.id) : undefined
    },
    {
      label: 'Critical',
      value: String(h.criticalInView),
      note: `${percentText(h.criticalInView, h.inView)} of these tickets`
    }
  ];

  return (
    <div className={styles.body}>
      <section aria-label="The story so far" className={styles.stories}>
        {story.map((s) => (
          <StoryCard key={s.kicker} story={s} />
        ))}
      </section>

      <section aria-label="Key figures" className={styles.kpis}>
        {kpis.map((k) => (
          <KpiTile key={k.label} {...k} />
        ))}
      </section>

      <div className={styles.row}>
        <ChartCard
          className={styles.trendWide}
          title={trendTitle}
          subtitle="When the blue line sits above the orange, the pile grows."
          empty={empty}
          legend={
            <>
              <LegendItem colour={PALETTE.raised} label="Raised" />
              <LegendItem colour={PALETTE.closed} label="Closed" />
            </>
          }
          chart={<TrendPlot buckets={t.buckets} kind="lines" label={`${trendTitle}: tickets raised and closed`} />}
          table={<TrendTable buckets={t.buckets} grain={t.grain} kind="lines" />}
        />
        <ChartCard
          className={styles.trendNarrow}
          title={openTitle}
          subtitle="The pile IT is carrying. Lower is better."
          empty={empty}
          chart={<TrendPlot buckets={t.buckets} kind="area" label={`${openTitle}: tickets still open`} />}
          table={<TrendTable buckets={t.buckets} grain={t.grain} kind="area" />}
        />
      </div>

      <div className={styles.row}>
        <ChartCard
          className={styles.wide}
          title="Who's holding what right now"
          subtitle="Open tickets per person, by where they are."
          empty={empty}
          legend={HOLDER_LEGEND.map((l) => (
            <LegendItem key={l.label} colour={l.colour} label={l.label} square={true} />
          ))}
          chart={
            team.length > 0 ? <HoldersBars holders={team} /> : <p className={styles.empty}>Nothing open right now.</p>
          }
          table={<HoldersTable holders={team} />}
        />
        <ChartCard
          className={styles.narrow}
          title="Closed this month"
          subtitle={`${monthName}'s leaderboard. Resets on the 1st.`}
          empty={empty}
          chart={<Leaderboard closers={board.closers} fastest={board.fastest} monthName={monthName} />}
          table={<LeaderboardTable closers={board.closers} />}
        />
      </div>

      <div className={styles.row}>
        <ChartCard
          className={styles.wide}
          title="What tickets are about"
          subtitle="Broken things vs requests, by area."
          empty={empty}
          legend={AREA_LEGEND.map((l) => (
            <LegendItem key={l.label} colour={l.colour} label={l.label} square={true} />
          ))}
          chart={<AreaBars areas={area.areas} uncategorised={area.uncategorised} />}
          table={<AreaTable areas={area.areas} uncategorised={area.uncategorised} />}
          footer={
            <p className={styles.quiet}>
              Older tickets saved under retired names (Software, Application, Access/Security) are counted with their area.
            </p>
          }
        />
        <ChartCard
          className={styles.narrow}
          title="How urgent they were"
          subtitle="Priority as it stands after IT's review."
          empty={empty}
          chart={<PriorityBar rows={pri} total={inView.length} />}
          table={<PriorityTable rows={pri} total={inView.length} />}
          footer={criticalNote !== undefined ? <p className={styles.quiet}>{criticalNote}</p> : undefined}
        />
      </div>

      <div className={styles.row}>
        <ChartCard
          className={styles.wide}
          title="When tickets arrive"
          subtitle="Darker = more tickets. Plan cover around the dark squares."
          empty={empty}
          chart={<HeatMap heat={heat} />}
          table={<HeatTable heat={heat} />}
          footer={<HeatLegend heat={heat} />}
        />
        <ChartCard
          className={styles.narrow}
          title="How long tickets take to close"
          subtitle="From raised to completed."
          empty={empty}
          chart={<SpeedBars rows={sp} />}
          table={<SpeedTable rows={sp} />}
          footer={
            h.medianDays !== undefined && h.p90Days !== undefined ? (
              <p className={styles.quiet}>
                Half close within <strong>{days1(h.medianDays)}</strong>. The slowest 1 in 10 take over{' '}
                <strong>{days1(h.p90Days)}</strong>.
              </p>
            ) : (
              <p className={styles.quiet}>Not enough closed tickets yet to say how long they take.</p>
            )
          }
        />
      </div>

      <div className={styles.row}>
        <ChartCard
          className={styles.wide}
          title="How old the open tickets are"
          subtitle={`${h.openNow} open. The right-hand end is where tickets go to be forgotten.`}
          empty={empty}
          chart={<AgeBars rows={ag} />}
          table={<AgeTable rows={ag} />}
          footer={
            oldBucket !== undefined && oldBucket.count > 0 ? (
              <div className={styles.alert} role="note">
                <AlertIcon />
                <span>
                  <strong>
                    {oldBucket.count === 1 ? '1 ticket has' : `${oldBucket.count} tickets have`} been open over 6 months.
                  </strong>
                  {h.oldestOpen !== undefined ? ` The oldest, ${h.oldestOpen.days} days.` : ''}
                </span>
              </div>
            ) : undefined
          }
        />
        <ChartCard
          className={styles.narrow}
          title="Gaps in the records"
          subtitle="What the charts can't see yet."
          empty={empty}
          chart={<GapBars gaps={gp} />}
          table={<GapTable gaps={gp} />}
        />
      </div>

      <p className={styles.footNote}>
        Each chart has a &quot;Show as table&quot; switch with the same numbers. Export gives the tickets behind every figure here.
      </p>
    </div>
  );
}
