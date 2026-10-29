// Who holds what, who closed what, what the tickets are about and how urgent
// they were. Each block has a plain-table twin for "Show as table".
import * as React from 'react';
import { AreaCount, Closer, days1, Holder } from '../stats/stats';
import { barPct, PALETTE, percentText } from './chartMath';
import styles from './Dashboard.module.scss';

const HOLD_COLOURS = { new: PALETTE.raised, progress: PALETTE.closed, hold: PALETTE.green };
const MEDALS = ['#c69214', '#8a8886', '#a5673f'];

// ---------------------------------------------------------------- who holds what

export function HoldersBars(props: { holders: Holder[] }): React.ReactElement {
  const { holders } = props;
  const top = holders.reduce((m, h) => Math.max(m, h.total), 0);
  return (
    <div className={styles.rows}>
      {holders.map((h) => (
        <div key={h.id} className={styles.holderRow}>
          <span className={styles.rowName} title={h.name}>
            {h.name}
          </span>
          <div className={styles.track} role="img" aria-label={`${h.new} new, ${h.inProgress} in progress, ${h.onHold} on hold`}>
            <span className={styles.seg} style={{ flexBasis: `${barPct(h.new, top)}%`, background: HOLD_COLOURS.new, minWidth: h.new > 0 ? 4 : 0 }} title={`${h.new} new`} />
            <span className={styles.seg} style={{ flexBasis: `${barPct(h.inProgress, top)}%`, background: HOLD_COLOURS.progress, minWidth: h.inProgress > 0 ? 4 : 0 }} title={`${h.inProgress} in progress`} />
            <span className={styles.seg} style={{ flexBasis: `${barPct(h.onHold, top)}%`, background: HOLD_COLOURS.hold, minWidth: h.onHold > 0 ? 4 : 0 }} title={`${h.onHold} on hold`} />
          </div>
          <span className={styles.rowNum}>{h.total}</span>
        </div>
      ))}
    </div>
  );
}

export function HoldersTable(props: { holders: Holder[] }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Person</th>
          <th scope="col">New</th>
          <th scope="col">In progress</th>
          <th scope="col">On hold</th>
          <th scope="col">Total</th>
        </tr>
      </thead>
      <tbody>
        {props.holders.map((h) => (
          <tr key={h.id}>
            <th scope="row">{h.name}</th>
            <td>{h.new}</td>
            <td>{h.inProgress}</td>
            <td>{h.onHold}</td>
            <td>{h.total}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export const HOLDER_LEGEND = [
  { colour: HOLD_COLOURS.new, label: 'New' },
  { colour: HOLD_COLOURS.progress, label: 'In progress' },
  { colour: HOLD_COLOURS.hold, label: 'On hold' }
];

// ---------------------------------------------------------------- closed this month

export interface LeaderboardProps {
  closers: Closer[];
  fastest?: Closer;
  monthName: string;
}

export function Leaderboard(props: LeaderboardProps): React.ReactElement {
  const { closers, fastest } = props;
  if (closers.length === 0) {
    return <p className={styles.empty}>Nothing closed yet this month — the board resets on the 1st.</p>;
  }
  const top = closers[0].closed;
  return (
    <div className={styles.boardWrap}>
      <ol className={styles.board}>
        {closers.map((c, i) => (
          <li key={c.id} className={styles.boardRow}>
            <span
              className={styles.medal}
              style={i < 3 ? { background: MEDALS[i], color: '#ffffff' } : undefined}
            >
              {i + 1}
            </span>
            <div className={styles.boardMid}>
              <span className={styles.rowName}>{c.name}</span>
              <div className={styles.thinTrack}>
                <div className={styles.thinFill} style={{ width: `${barPct(c.closed, top)}%` }} />
              </div>
            </div>
            <span className={styles.boardNum}>{c.closed}</span>
          </li>
        ))}
      </ol>
      {fastest !== undefined && fastest.medianDays !== undefined && (
        <p className={styles.callout}>
          Fastest median close this month: <strong>{`${fastest.name}, ${days1(fastest.medianDays)}`}</strong>.
        </p>
      )}
    </div>
  );
}

export function LeaderboardTable(props: { closers: Closer[] }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Rank</th>
          <th scope="col">Name</th>
          <th scope="col">Closed</th>
          <th scope="col">Median days</th>
        </tr>
      </thead>
      <tbody>
        {props.closers.map((c, i) => (
          <tr key={c.id}>
            <td>{i + 1}</td>
            <th scope="row">{c.name}</th>
            <td>{c.closed}</td>
            <td>{c.medianDays !== undefined ? days1(c.medianDays) : '—'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ---------------------------------------------------------------- what tickets are about

export function AreaBars(props: { areas: AreaCount[]; uncategorised: number }): React.ReactElement {
  const { areas, uncategorised } = props;
  const top = areas.reduce((m, a) => Math.max(m, a.total), 0);
  return (
    <div className={styles.areaBlock}>
      <div className={styles.rows}>
        {areas.map((a) => (
          <div key={a.area} className={styles.areaRow}>
            <span className={styles.rowName}>{a.label}</span>
            <div className={styles.stack}>
              <span className={styles.seg} style={{ flexBasis: `${barPct(a.incident, top)}%`, background: PALETTE.raised, minWidth: a.incident > 0 ? 4 : 0 }} title={`${a.incident} something's broken`} />
              <span className={styles.seg} style={{ flexBasis: `${barPct(a.request, top)}%`, background: PALETTE.closed, minWidth: a.request > 0 ? 4 : 0 }} title={`${a.request} requests`} />
              <span className={styles.seg} style={{ flexBasis: `${barPct(a.other, top)}%`, background: PALETTE.other, minWidth: a.other > 0 ? 4 : 0 }} title={`${a.other} other`} />
            </div>
            <span className={styles.rowNum}>{a.total}</span>
          </div>
        ))}
      </div>
      {uncategorised > 0 && <p className={styles.quiet}>{`No category: ${uncategorised}`}</p>}
    </div>
  );
}

export function AreaTable(props: { areas: AreaCount[]; uncategorised: number }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Area</th>
          <th scope="col">Something&apos;s broken</th>
          <th scope="col">A request</th>
          <th scope="col">Other kinds</th>
          <th scope="col">Total</th>
        </tr>
      </thead>
      <tbody>
        {props.areas.map((a) => (
          <tr key={a.area}>
            <th scope="row">{a.label}</th>
            <td>{a.incident}</td>
            <td>{a.request}</td>
            <td>{a.other}</td>
            <td>{a.total}</td>
          </tr>
        ))}
        {props.uncategorised > 0 && (
          <tr>
            <th scope="row">No category</th>
            <td colSpan={3}>—</td>
            <td>{props.uncategorised}</td>
          </tr>
        )}
      </tbody>
    </table>
  );
}

export const AREA_LEGEND = [
  { colour: PALETTE.raised, label: "Something's broken" },
  { colour: PALETTE.closed, label: 'A request' },
  { colour: PALETTE.other, label: 'Other kinds' }
];

// ---------------------------------------------------------------- how urgent

export const PRIORITY_COLOURS: Record<string, string> = {
  Critical: PALETTE.critical,
  High: PALETTE.alert,
  Normal: PALETTE.raised,
  Low: PALETTE.calm,
  'Not set': PALETTE.notSet
};

export interface PriorityRow {
  priority: string;
  count: number;
}

/** Every priority in order, zeros included, so the legend never changes shape. */
export function priorityRows(list: { priority: string; count: number }[]): PriorityRow[] {
  return Object.keys(PRIORITY_COLOURS).map((p) => ({
    priority: p,
    count: list.reduce((sum, x) => (x.priority === p ? sum + x.count : sum), 0)
  }));
}

export function PriorityBar(props: { rows: PriorityRow[]; total: number }): React.ReactElement {
  const { rows, total } = props;
  return (
    <div className={styles.priorityBlock}>
      <div className={styles.priorityBar} role="img" aria-label="Share of tickets by priority">
        {rows.map((r) => (
          <span
            key={r.priority}
            className={styles.priSeg}
            style={{ flexGrow: r.count, background: PRIORITY_COLOURS[r.priority] }}
            title={`${r.priority}: ${r.count} tickets`}
          />
        ))}
      </div>
      <ul className={styles.priorityLegend}>
        {rows.map((r) => (
          <li key={r.priority} className={styles.priItem}>
            <span className={styles.priSwatch} style={{ background: PRIORITY_COLOURS[r.priority] }} aria-hidden="true" />
            <span>{r.priority}</span>
            <strong className={styles.priPct}>{percentText(r.count, total)}</strong>
          </li>
        ))}
      </ul>
    </div>
  );
}

export function PriorityTable(props: { rows: PriorityRow[]; total: number }): React.ReactElement {
  return (
    <table className={styles.table}>
      <thead>
        <tr>
          <th scope="col">Priority</th>
          <th scope="col">Tickets</th>
          <th scope="col">Share</th>
        </tr>
      </thead>
      <tbody>
        {props.rows.map((r) => (
          <tr key={r.priority}>
            <th scope="row">{r.priority}</th>
            <td>{r.count}</td>
            <td>{percentText(r.count, props.total)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
