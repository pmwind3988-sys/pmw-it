// Who holds what, who closed what, what the tickets are about and how urgent
// they were. Each block has a plain-table twin for "Show as table".
import * as React from 'react';
import { Picker } from '../stats/drill';
import { AreaCount, AreaPart, Closer, days1, Holder, HoldState } from '../stats/stats';
import { barPct, PALETTE, percentText, ticketsText } from './chartMath';
import { Mark, Segment } from './Marks';
import styles from './Dashboard.module.scss';

const HOLD_COLOURS = { new: PALETTE.raised, progress: PALETTE.closed, hold: PALETTE.green };
const MEDALS = ['#c69214', '#8a8886', '#a5673f'];

// ---------------------------------------------------------------- who holds what

const pickingClass = (pick: Picker): string => (pick.picked !== undefined ? styles.picking : '');

export function HoldersBars(props: { holders: Holder[]; pick: Picker }): React.ReactElement {
  const { holders, pick } = props;
  const top = holders.reduce((m, h) => Math.max(m, h.total), 0);
  const parts: { state: HoldState; count: (h: Holder) => number; colour: string; words: string }[] = [
    { state: 'new', count: (h) => h.new, colour: HOLD_COLOURS.new, words: 'new' },
    { state: 'inProgress', count: (h) => h.inProgress, colour: HOLD_COLOURS.progress, words: 'in progress' },
    { state: 'onHold', count: (h) => h.onHold, colour: HOLD_COLOURS.hold, words: 'on hold' }
  ];
  return (
    <div className={`${styles.rows} ${pickingClass(pick)}`}>
      {holders.map((h) => (
        <div key={h.id} className={styles.holderRow}>
          <Mark
            pick={pick}
            drill={{ kind: 'holder', id: h.id }}
            label={`${h.name}: ${ticketsText(h.total)} open`}
            className={`${styles.rowName} ${styles.markRow}`}
          >
            {h.name}
          </Mark>
          <div className={styles.track}>
            {parts.map((p) => (
              <Segment
                key={p.state}
                pick={pick}
                drill={{ kind: 'holder', id: h.id, state: p.state }}
                count={p.count(h)}
                label={`${h.name}: ${p.count(h)} ${p.words}`}
                className={styles.seg}
                style={{ flexBasis: `${barPct(p.count(h), top)}%`, background: p.colour, minWidth: p.count(h) > 0 ? 4 : 0 }}
              />
            ))}
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
  pick: Picker;
}

export function Leaderboard(props: LeaderboardProps): React.ReactElement {
  const { closers, fastest, pick } = props;
  if (closers.length === 0) {
    return <p className={styles.empty}>Nothing closed yet this month — the board resets on the 1st.</p>;
  }
  const top = closers[0].closed;
  return (
    <div className={styles.boardWrap}>
      <ol className={`${styles.board} ${pickingClass(pick)}`}>
        {closers.map((c, i) => (
          <li key={c.id}>
            <Mark
              pick={pick}
              drill={{ kind: 'closer', id: c.id }}
              label={`${c.name}: ${ticketsText(c.closed)} closed this month`}
              className={`${styles.boardRow} ${styles.markRow}`}
            >
              <span
                className={styles.medal}
                style={i < 3 ? { background: MEDALS[i], color: '#ffffff' } : undefined}
              >
                {i + 1}
              </span>
              <span className={styles.boardMid}>
                <span className={styles.rowName}>{c.name}</span>
                <span className={styles.thinTrack}>
                  <span className={styles.thinFill} style={{ width: `${barPct(c.closed, top)}%` }} />
                </span>
              </span>
              <span className={styles.boardNum}>{c.closed}</span>
            </Mark>
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

const AREA_PARTS: { part: AreaPart; colour: string; words: string }[] = [
  { part: 'incident', colour: PALETTE.raised, words: "something's broken" },
  { part: 'request', colour: PALETTE.closed, words: 'requests' },
  { part: 'other', colour: PALETTE.other, words: 'other' }
];

export function AreaBars(props: { areas: AreaCount[]; uncategorised: number; pick: Picker }): React.ReactElement {
  const { areas, uncategorised, pick } = props;
  const top = areas.reduce((m, a) => Math.max(m, a.total), 0);
  return (
    <div className={`${styles.areaBlock} ${pickingClass(pick)}`}>
      <div className={styles.rows}>
        {areas.map((a) => (
          <div key={a.area} className={styles.areaRow}>
            <Mark
              pick={pick}
              drill={{ kind: 'area', area: a.area }}
              label={`${a.label}: ${ticketsText(a.total)}`}
              className={`${styles.rowName} ${styles.markRow}`}
            >
              {a.label}
            </Mark>
            <div className={styles.stack}>
              {AREA_PARTS.map((p) => (
                <Segment
                  key={p.part}
                  pick={pick}
                  drill={{ kind: 'area', area: a.area, part: p.part }}
                  count={a[p.part]}
                  label={`${a.label}: ${a[p.part]} ${p.words}`}
                  className={styles.seg}
                  style={{ flexBasis: `${barPct(a[p.part], top)}%`, background: p.colour, minWidth: a[p.part] > 0 ? 4 : 0 }}
                />
              ))}
            </div>
            <span className={styles.rowNum}>{a.total}</span>
          </div>
        ))}
      </div>
      {uncategorised > 0 && (
        <Mark
          pick={pick}
          drill={{ kind: 'area', area: '' }}
          label={`No category: ${ticketsText(uncategorised)}`}
          className={`${styles.quiet} ${styles.gapButton}`}
        >
          {`No category: ${uncategorised}`}
        </Mark>
      )}
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

export function PriorityBar(props: { rows: PriorityRow[]; total: number; pick: Picker }): React.ReactElement {
  const { rows, total, pick } = props;
  return (
    <div className={`${styles.priorityBlock} ${pickingClass(pick)}`}>
      <div className={styles.priorityBar}>
        {rows.map((r) => (
          <Segment
            key={r.priority}
            pick={pick}
            drill={{ kind: 'priority', priority: r.priority }}
            count={r.count}
            label={`${r.priority}: ${ticketsText(r.count)}`}
            className={styles.priSeg}
            style={{ flexGrow: r.count, background: PRIORITY_COLOURS[r.priority] }}
          />
        ))}
      </div>
      <ul className={styles.priorityLegend}>
        {rows.map((r) => (
          <li key={r.priority}>
            {r.count > 0 ? (
              <Mark
                pick={pick}
                drill={{ kind: 'priority', priority: r.priority }}
                label={`${r.priority}: ${ticketsText(r.count)}`}
                className={`${styles.priItem} ${styles.markRow}`}
              >
                <span className={styles.priSwatch} style={{ background: PRIORITY_COLOURS[r.priority] }} aria-hidden="true" />
                <span>{r.priority}</span>
                <strong className={styles.priPct}>{percentText(r.count, total)}</strong>
              </Mark>
            ) : (
              <span className={styles.priItem}>
                <span className={styles.priSwatch} style={{ background: PRIORITY_COLOURS[r.priority] }} aria-hidden="true" />
                <span>{r.priority}</span>
                <strong className={styles.priPct}>{percentText(r.count, total)}</strong>
              </span>
            )}
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
